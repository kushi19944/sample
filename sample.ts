import RPA from "ts-rpa";
require('dotenv').config();
const test = process.env.CyNumber1;
RPA.Logger.info(test);

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
