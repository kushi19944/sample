import RPA from 'ts-rpa';
import { WebDriver, By, FileDetector, Key } from 'selenium-webdriver';

async function test() {
  await RPA.WebBrowser.get('https://github.com/kushi19944');
  await RPA.sleep(10000);
  const C = await RPA.WebBrowser.findElementByXPath(
    '/html/body/div[1]/header/div[3]/div/div/form/label/input[1]'
  );
  await RPA.WebBrowser.sendKeys(C, ['hello']);
  await RPA.WebBrowser.sendKeys(C, [Key.ENTER]);
  await RPA.WebBrowser.takeScreenshot();
}
test();
