const puppeteer = require("puppeteer");

const startBrowser = async () => {
  let browser;
  try {
    browser = await puppeteer.launch({
      timeout: 60 * 60000,
      headless: false,
      args: ["--disable-setuid-sandbox"],
      ignoreHTTPSErrors: true,
    });
  } catch (error) {
    console.log("không tạo được browser:" + error);
  }
  return browser;
};

module.exports = startBrowser;
