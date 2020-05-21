const webdriver = require("selenium-webdriver");

function buildCapabilities() {
  switch (process.env.BROWSER) {
    case "ie": {
      process.env.PATH = `${process.env.PATH};${__dirname}/Selenium.WebDriver.IEDriver.3.150.0/driver/;`;
      const capabilities = webdriver.Capabilities.ie();
      capabilities.set("ignoreProtectedModeSettings", true);
      capabilities.set("ignoreZoomSetting", true);
      return capabilities;
    }
    case "safari": {
      return webdriver.Capabilities.safari();
    }
  }
}

async function main() {
  let driver;
  try {
    const capabilities = await buildCapabilities();
    driver = await new webdriver.Builder()
      .withCapabilities(capabilities)
      .build();

    await driver.get("https://www.youtube.com/");
    const title = await driver.getTitle();
    console.log("title", title);
  } catch (e) {
    console.error(e);
    process.exit(1);
  }
  driver && (await driver.quit());
}

main();
