import { SPHttpClient } from "@microsoft/sp-http";

// ─── SharePoint List Service ──────────────────────────────────────────────────
// Auto-discovers all visible lists on the site and reads items in user context.

// spListService.ts – helper to compose select/expand
type FieldDef = {
  InternalName: string;
  Title: string;
  FieldTypeKind: number;        // 7=Lookup, 20=User, etc.
  LookupField?: string;         // e.g., 'Title' by default
  AllowMultipleValues?: boolean;
};

function buildSelectExpand(
  fields: FieldDef[],
  requestedInternalNames: string[]
): { select: string[]; expand: string[] } {
  const select: string[] = ["Id"];         // always include Id
  const expand = new Set<string>();

  const byName = new Map(fields.map(f => [f.InternalName, f]));

  for (const name of requestedInternalNames) {
    const f = byName.get(name);
    if (!f) { select.push(name); continue; }

    switch (f.FieldTypeKind) {
      case 20: { // User
        select.push(`${name}/Title`, `${name}/EMail`, `${name}/Id`);
        expand.add(name);
        break;
      }
      case 7: {  // Lookup
        const lookupCol = f.LookupField || "Title";
        select.push(`${name}/Id`, `${name}/${lookupCol}`);
        expand.add(name);
        break;
      }
      default: {
        // Text/Number/Date/Choice/Note/URL/etc.
        select.push(name);
        break;
      }
    }
  }

  return { select, expand: Array.from(expand) };
}


export interface ListInfo {
  Id: string;
  Title: string;
  ItemCount: number;
  Description?: string;
}

export interface ListQueryResult {
  listTitle: string;
  columns: string[];
  items: Record<string, unknown>[];
  totalCount: number;
  truncated: boolean;
}

// ─── Discover all lists ───────────────────────────────────────────────────────

export async function getAllLists(
  spHttpClient: SPHttpClient,
  webUrl: string
): Promise<ListInfo[]> {
  try {
    const url =
      `${webUrl}/_api/web/lists` +
      `?$select=Id,Title,ItemCount,Description` +
      `&$filter=Hidden eq false and BaseType eq 0` +
      `&$orderby=Title` +
      `&$top=50`;

    const res = await spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!res.ok) return [];

    const data = await res.json() as { value: ListInfo[] };
    return data.value ?? [];
  } catch {
    return [];
  }
}

// ─── Fuzzy match list by name hint ───────────────────────────────────────────

export function findListByHint(
  lists: ListInfo[],
  hint: string
): ListInfo | null {
  if (!hint || lists.length === 0) return null;

  const h = hint.toLowerCase().trim();

  // 1 — exact
  const exact = lists.find((l) => l.Title.toLowerCase() === h);
  if (exact) return exact;

  // 2 — starts with
  const starts = lists.find((l) => l.Title.toLowerCase().startsWith(h));
  if (starts) return starts;

  // 3 — hint contains list name or list name contains hint
  const contains = lists.find(
    (l) =>
      l.Title.toLowerCase().includes(h) ||
      h.includes(l.Title.toLowerCase())
  );
  if (contains) return contains;

  // 4 — significant words match
  const words = h.split(/\s+/).filter((w) => w.length > 2);
  const wordMatch = lists.find((l) => {
    const t = l.Title.toLowerCase();
    return words.filter((w) => t.includes(w)).length >= Math.max(1, Math.floor(words.length * 0.6));
  });

  return wordMatch ?? null;
}

// ─── Fetch list items ─────────────────────────────────────────────────────────

export async function getListItems(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  listTitle: string,
  requested: string[],      // e.g., ["Title","Employee","ActiveProject_x0028_s_x0029_","ShiftTimings",...]
  top = 100
) {
  // 1) Get schema
  const fieldsRes = await spHttpClient.get(
    `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/fields` +
    "?$select=InternalName,Title,FieldTypeKind,LookupField,AllowMultipleValues,Hidden,ReadOnlyField" +
    "&$filter=Hidden eq false and ReadOnlyField eq false",
    SPHttpClient.configurations.v1
  );
  const fieldsJson = await fieldsRes.json();
  const fields: FieldDef[] = fieldsJson.value;

  // 2) Compose select/expand
  const { select, expand } = buildSelectExpand(fields, requested);
  const selectStr = `$select=${encodeURIComponent(select.join(","))}`;
  const expandStr = expand.length ? `&$expand=${expand.join(",")}` : "";
  const url =
    `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items` +
    `?${selectStr}${expandStr}&$top=${top}&$orderby=Modified desc`;

  // 3) Fetch items
  const itemsRes = await spHttpClient.get(url, SPHttpClient.configurations.v1);
  if (!itemsRes.ok) {
    const errText = await itemsRes.text().catch(() => "");
    throw new Error(`Failed to read items: ${itemsRes.status} ${errText.slice(0, 300)}`);
  }
  const itemsJson = await itemsRes.json();
  return { listTitle, columns: requested, items: itemsJson.value as Record<string, unknown>[] };
}

