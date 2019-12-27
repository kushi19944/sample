import RPA from "ts-rpa";

(async () => {
  try {
    await RPA.WebBrowser.get("https://github.com/kushi19944/sample/");
    await RPA.sleep(500);
    await RPA.WebBrowser.takeScreenshot();
  } catch (error) {
    RPA.SystemLogger.error(error);
  } finally {
    await RPA.WebBrowser.quit();
  }
})();
