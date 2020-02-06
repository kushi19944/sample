import RPA from 'ts-rpa';
import { WebDriver, By, FileDetector, Key } from 'selenium-webdriver';

async function test() {
  await RPA.WebBrowser.get('https://www.google.com/?hl=ja');
  await RPA.sleep(10000);
  const C = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="tsf"]/div[2]/div[1]/div[1]/div/div[2]/input'
  );
  await RPA.WebBrowser.sendKeys(C, ['hello']);
  await RPA.WebBrowser.sendKeys(C, [Key.ENTER]);
  await RPA.WebBrowser.takeScreenshot();
}
test();
