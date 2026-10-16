// Accesso amministratore: POST /api/login  { password }
import { checkPassword, makeToken, json, fail } from "../lib/common.mjs";

export default async req => {
  try {
    if (req.method !== "POST") return json({ error: "Metodo non consentito" }, 405);
    const { password } = await req.json().catch(() => ({}));
    if (!checkPassword(password)) {
      await new Promise(r => setTimeout(r, 800));   // rallenta i tentativi
      return json({ error: "Password errata." }, 401);
    }
    return json({ token: makeToken() });
  } catch (e) { return fail(e); }
};
export const config = { path: "/api/login" };
