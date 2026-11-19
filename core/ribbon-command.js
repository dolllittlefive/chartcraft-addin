// Function commands share the pane's operation guard and instance lifecycle.
// Always acknowledge Office, including validation failures and busy rejection.
export function registerCreateCommand({office, action, ready, create, reportError}) {
  office.actions.associate(action, async event => {
    try {
      await ready;
      await create();
    } catch (error) {
      await reportError(error);
    } finally {
      event.completed();
    }
  });
}
