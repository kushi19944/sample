import RPA from "ts-rpa";
require('dotenv').config();
const test = process.env.CyNumber;
RPA.Logger.info(test);
const test2 = process.env.Youtube_Card_Setting_SheetID;
RPA.Logger.info(test2);

(async () => {
  try {
    await RPA.WebBrowser.get("https://www.google.com/");
    await RPA.sleep(5000);
    await RPA.WebBrowser.takeScreenshot();
  } catch (error) {
    RPA.SystemLogger.error(error);
  } finally {
    await RPA.WebBrowser.quit();
  }
})();
