// Turning Microsoft Graph's errors into something an employee can act on.
//
// The app writes to SharePoint as the signed-in user (delegated permissions),
// so a refused write comes back as a bare 403 "Access denied" — which says
// nothing about what is wrong or who can fix it. It means that account has no
// edit rights on the expenses site, and only a site owner can grant them.

export const NO_ACCESS_TEXT =
  "You're signed in with your Arqus account, but it doesn't have edit access " +
  "to the expenses SharePoint site, so nothing can be filed with it yet. " +
  "Ask Marnix, Stijn or Anton to give your account access to the site, then " +
  "reload this page.";

export const isAccessDenied = (err) =>
  err?.status === 403 || err?.code === "accessDenied";

export function explainError(err) {
  if (isAccessDenied(err)) return NO_ACCESS_TEXT;
  if (err?.status === 401) return "Your sign-in expired — sign out and back in.";
  return err?.message || "Something went wrong.";
}

// The raw Graph error, kept visible in small print so a founder debugging
// someone's access has the code and status to hand.
export function errorDetail(err) {
  const bits = [err?.code, err?.status ? `HTTP ${err.status}` : null].filter(Boolean);
  return `Microsoft Graph said: ${err?.message || "Access denied"}` +
    (bits.length ? ` (${bits.join(", ")})` : "");
}
