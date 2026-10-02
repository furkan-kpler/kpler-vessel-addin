/**
 * Kpler Vessel API CORS proxy.
 *
 * Problem this solves: the Kpler GraphQL endpoint does not return
 * Access-Control-Allow-Origin on its preflight response, so any browser-based
 * caller (like this Excel add-in's task pane) gets blocked before the real
 * request even goes out - curl works fine because curl doesn't enforce CORS,
 * but a browser does.
 *
 * What this worker does and does NOT do:
 * - It does NOT hold or inject any credential of its own. The customer's token
 *   still travels from the add-in's Connection settings, in the Authorization
 *   header, exactly as they typed it. This worker only forwards that header
 *   through to the real endpoint and adds permissive CORS headers to the
 *   response so the browser will accept it.
 * - It is a pure pass-through: same request body, same auth, same response -
 *   just with CORS unblocked. No logging, no credential storage.
 */

const KPLER_ENDPOINT_DEFAULT = "https://api.kpler.com/v2/maritime/vessels/graphql";

function corsHeaders() {
  return {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization",
    "Access-Control-Max-Age": "86400",
  };
}

function jsonError(message, status) {
  return new Response(JSON.stringify({ errors: [{ message }] }), {
    status,
    headers: { "Content-Type": "application/json", ...corsHeaders() },
  });
}

export default {
  async fetch(request, env) {
    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204, headers: corsHeaders() });
    }

    if (request.method !== "POST") {
      return jsonError("Method not allowed - POST only.", 405);
    }

    const auth = request.headers.get("Authorization");
    if (!auth) {
      return jsonError("Missing Authorization header.", 401);
    }

    let body;
    try {
      body = await request.text();
    } catch (err) {
      return jsonError(`Could not read request body: ${err.message}`, 400);
    }

    const target = (env && env.KPLER_ENDPOINT) || KPLER_ENDPOINT_DEFAULT;

    let upstream;
    try {
      upstream = await fetch(target, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: auth,
        },
        body,
      });
    } catch (err) {
      return jsonError(`Upstream request to Kpler failed: ${err.message}`, 502);
    }

    const responseBody = await upstream.text();
    return new Response(responseBody, {
      status: upstream.status,
      headers: {
        "Content-Type": upstream.headers.get("Content-Type") || "application/json",
        ...corsHeaders(),
      },
    });
  },
};
