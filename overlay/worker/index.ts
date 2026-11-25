import handler from "vinext/server/fetch-handler";
import { handleMcp } from "../lib/mcp.mjs";
import { handleTransfer } from "../lib/transfer.mjs";
import { handleAuth } from "../lib/ms-auth.mjs";
import { graphClient } from "../lib/graph.mjs";

export default {
  async fetch(request: Request, env: Cloudflare.Env, ctx: ExecutionContext) {
    const url = new URL(request.url);
    const user = request.headers.get("oai-authenticated-user-id");
    if (url.pathname === "/mcp") {
      return handleMcp(request, env, user, graphClient(env, user, true));
    }
    if (url.pathname.startsWith("/api/connect/")) return handleAuth(request, env, user);
    if (url.pathname.startsWith("/api/files/")) {
      if (!user) return Response.json({error:"Sign in to your private OneDrive plugin."}, {status:401});
      if (!["GET","HEAD"].includes(request.method) && request.headers.get("Origin") !== url.origin) {
        return Response.json({error:"Open the private transfer page to make changes."}, {status:403});
      }
      return handleTransfer(request, env, user, graphClient(env, user));
    }
    return handler.fetch(request, env, ctx);
  },
};
