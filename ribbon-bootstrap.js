// Office can dispatch a function command before ES modules have loaded.
// Register synchronously; the pane resolves readiness once its controls exist.
(() => {
  const routes = JSON.parse(document.getElementById('cc-ribbon-routes').textContent);
  let resolve;
  const ready = new Promise(done => { resolve = done; });
  window.CCResolveRibbon = resolve;
  for (const route of routes) {
    const handler = async event => {
      try {
        const run = await ready;
        await run(route);
      } catch (error) {
        // The pane reports user-facing errors; acknowledge Office even if its
        // error UI cannot be shown, so a subsequent click is not left queued.
        console.error(error);
      } finally {
        event.completed();
      }
    };
    window[route.action] = handler;
    Office.actions.associate(route.action, handler);
  }
})();
