import RPA from "ts-rpa";

(async () => {
  try {
    await RPA.WebBrowser.get("https://www.google.com/");
    await RPA.sleep(15000);
    await RPA.WebBrowser.takeScreenshot();
  } catch (error) {
    RPA.SystemLogger.error(error);
  } finally {
    await RPA.WebBrowser.quit();
  }
})();
