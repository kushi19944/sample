import RPA from "ts-rpa";

(async () => {
  try {
    await RPA.WebBrowser.get("https://www.google.com/");
    const ele = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'//*[@id="hptl"]/a[2]'}),5000);
    const text = await ele.getText();
    if(text == "ストア"){
     await RPA.WebBrowser.mouseClick(ele); 
    }
    await RPA.sleep(15000);
    await RPA.WebBrowser.takeScreenshot();
  } catch (error) {
    RPA.SystemLogger.error(error);
  } finally {
    await RPA.WebBrowser.quit();
  }
})();
