import { MSGraphClientFactory } from "@microsoft/sp-http";

// ─── Graph Client Wrapper ────────────────────────────────────────────────────
// Wraps MSGraphClientFactory to produce typed get/post helpers.
// All requests are made in the delegated context of the current user.

let _factory: MSGraphClientFactory | null = null;

export function initGraphClientFactory(factory: MSGraphClientFactory): void {
  _factory = factory;
}

async function getClient(): Promise<any> {
  if (!_factory) {
    throw new Error("GraphClientFactory not initialized. Call initGraphClientFactory first.");
  }
  // "3" = Microsoft Graph v1.0
  return _factory.getClient("3");
}

export async function graphGet<T = unknown>(apiPath: string): Promise<T> {
  const client = await getClient();
  return client.api(apiPath).version("v1.0").get();
}

export async function graphPost<T = unknown>(
  apiPath: string,
  body: unknown
): Promise<T> {
  const client = await getClient();
  return client
    .api(apiPath)
    .version("v1.0")
    .header("Content-Type", "application/json")
    .post(body);
}
