/* global Office */
// No ribbon actions call ExecuteFunction in this add-in (the only button opens the
// task pane directly), so this file has nothing to register. It exists because the
// manifest's FunctionFile points here for older Office hosts that expect one.
Office.onReady();
