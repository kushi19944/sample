import RPA from "ts-rpa";

(async () => {
  try {
    await RPA.WebBrowser.get("https://www.google.com/");
    const input = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'//*[@id="tsf"]/div[2]/div[1]/div[1]/div/div[2]/input'}),5000);
    await RPA.WebBrowser.sendKeys(input,[`${google}`]);
    const ele = await RPA.WebBrowser.findElementByXPath('//*[@id="hptl"]/a[2]');
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
