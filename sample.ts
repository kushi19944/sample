import RPA from "ts-rpa";

(async () => {
  try {
    await RPA.WebBrowser.get("https://github.com/kushi19944/sample/");
    for(let i = 0; i < 30; i++){
     RPA.Logger.log(i);
     RPA.sleep(500);
    }
    await RPA.WebBrowser.takeScreenshot();
  } catch (error) {
    RPA.SystemLogger.error(error);
  } finally {
    await RPA.WebBrowser.quit();
  }
})();
