import { graphGet } from "./graphClient";
import type { UserProfile, ManagerProfile } from "./types";
import { trackError } from "../utils/telemetry";

// ─── Graph User Service ───────────────────────────────────────────────────────

interface GraphUser {
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
  jobTitle?: string;
  department?: string;
}

export async function getMe(): Promise<UserProfile> {
  try {
    const me = await graphGet<GraphUser>(
      "/me?$select=displayName,mail,userPrincipalName,jobTitle,department"
    );
    return {
      displayName: me.displayName ?? "Unknown User",
      mail: me.mail ?? me.userPrincipalName ?? "",
      jobTitle: me.jobTitle,
      department: me.department,
    };
  } catch (error) {
    trackError("graphUserService.getMe", error);
    throw error;
  }
}

export async function getManager(): Promise<ManagerProfile | null> {
  try {
    const mgr = await graphGet<GraphUser>(
      "/me/manager?$select=displayName,mail,userPrincipalName"
    );
    return {
      displayName: mgr.displayName ?? "Your Manager",
      mail: mgr.mail ?? mgr.userPrincipalName ?? "",
    };
  } catch (error) {
    // 404 = no manager configured, not a hard error
    const status = (error as { statusCode?: number }).statusCode;
    if (status === 404) return null;
    trackError("graphUserService.getManager", error);
    return null; // return null rather than throwing — manager is optional
  }
}
