(() => {
  "use strict";
  let lastResult = "";
  const origin = new URL(document.referrer || window.location.href).origin;
  const send = (type, data) => window.parent.postMessage({ isStreamlitMessage: true, type, ...data }, origin);
  const read = name => {
    const prefix = encodeURIComponent(name) + "=";
    const entry = document.cookie.split(";").map(value => value.trim()).find(value => value.startsWith(prefix));
    return entry ? decodeURIComponent(entry.slice(prefix.length)) : null;
  };
  window.addEventListener("message", event => {
    if (event.source !== window.parent || event.origin !== origin || event.data?.type !== "streamlit:render") return;
    const args = event.data.args;
    let error = null;
    let token = null;
    try {
      const secure = window.location.protocol === "https:" ? "; Secure" : "";
      const cookie = encodeURIComponent(args.cookie_name);
      if (args.action === "set") {
        document.cookie = `${cookie}=${encodeURIComponent(args.token)}; Path=/; Max-Age=${args.max_age}; SameSite=Lax${secure}`;
        if (read(args.cookie_name) !== args.token) error = "The browser did not accept the login cookie.";
      } else if (args.action === "clear") {
        document.cookie = `${cookie}=; Path=/; Max-Age=0; SameSite=Lax${secure}`;
      }
      token = read(args.cookie_name);
    } catch (_) {
      error = "The browser could not access the login cookie.";
    }
    const result = { ready: true, token, operation_id: args.operation_id, error };
    const serialized = JSON.stringify(result);
    if (serialized !== lastResult) {
      lastResult = serialized;
      send("streamlit:setComponentValue", { value: result, dataType: "json" });
    }
    send("streamlit:setFrameHeight", { height: 0 });
  });
  send("streamlit:componentReady", { apiVersion: 1 });
})();
