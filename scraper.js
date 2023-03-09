const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");
const ac = require("@antiadmin/anticaptchaofficial");
ac.setAPIKey(
  process.env.ANTI_CAPTCHA_KEY || "15142e9e43076b50f2c78247a94129ca"
);
const scrapeCategory = async (browser, url) =>
  new Promise(async (resolve, reject) => {
    try {
      const arrayA = [];
      const timestamp = new Date().getTime();
      let page = await browser.newPage();
      console.log(">> Mở tab mới...");

      await page.goto(url);
      const client = await page.target().createCDPSession();
      await client.send("Page.setDownloadBehavior", {
        behavior: "allow",
        downloadPath: path.resolve("./filetax"),
      });
      await page.click(
        "#bodyP > div.khungtong > div.frm_login > div.khungbaolongin > div.left-conten > div > div:nth-child(2) > a"
      );

      await page.waitForTimeout(2000);
      await page.evaluate(() => {
        document
          .querySelector(
            "#bodyP > div > div.banner > div.dangnhap > span:nth-child(2) > button"
          )
          .click();
      });

      await page.waitForSelector("#_userName");
      await page.focus("#_userName");
      await page.type("#_userName", "0315375348-QL");

      await page.waitForSelector("#password");
      await page.focus("#password");
      await page.type("#password", "@Gonext$12");

      const myElement = await page.$(
        "#loginForm > table > tbody > tr:nth-child(4) > td > div > div.hien_mxn"
      );

      const screenshotBuffer = await myElement.screenshot();
      fs.writeFileSync(
        path.resolve(__dirname, `${timestamp}-captcha.png`),
        screenshotBuffer
      );
      const captcha = fs.readFileSync(
        path.resolve(__dirname, `${timestamp}-captcha.png`),
        {
          encoding: "base64",
        }
      );

      let valueCaptcha = null;
      for (let retryIndex = 0; retryIndex < 1000; retryIndex++) {
        try {
          // eslint-disable-next-line no-await-in-loop
          valueCaptcha = await ac.solveImage(captcha, true);

          break;
        } catch (error) {
          // eslint-disable-next-line no-continue
          continue;
        }
      }
      fs.rmSync(path.resolve(__dirname, `${timestamp}-captcha.png`), {
        recursive: true,
      });

      console.log(valueCaptcha);
      await page.waitForSelector("#vcode");
      await page.focus("#vcode");
      await page.type("#vcode", valueCaptcha);

      await page.waitForSelector("#dangnhap");
      await page.evaluate(() => {
        document.querySelector("#dangnhap").click();
      });

      await page.waitForNavigation();

      await page.waitForSelector("#tabmenu > li.li-3 > a");

      await page.click("#tabmenu > li.li-3 > a");

      await page.waitForSelector("#sc3 > ul > li:nth-child(8) > a");

      await page.evaluate(() => {
        document.querySelector("#sc3 > ul > li:nth-child(8) > a").click();
      });

      await page.waitForTimeout(5000);
      // await page.click(".button_vuong");
      await page.evaluate(() => {
        document
          .querySelector("#tranFrame")
          .contentWindow.document.body.querySelector(".button_vuong")
          .click();
        return null;
      });

      await page.waitForTimeout(5000);

      const number = await page.evaluate(() => {
        return document
          .querySelector("#tranFrame")
          .contentWindow.document.body.querySelector(
            `#currAcc > b:nth-child(7)`
          )
          .innerHTML.trim();
      });
      const numberPage = parseInt(number, 10);
      await page.waitForTimeout(2000);

      for (let j = 1; j <= numberPage; j++) {
        if (j === 1) {
          await page.waitForTimeout(10000);
        } else if (j === 2) {
          await page.waitForTimeout(10000);
          await page.evaluate(() => {
            document
              .querySelector("#tranFrame")
              .contentWindow.document.body.querySelector(
                "#currAcc > a:nth-child(1)"
              )
              .click();
            return null;
          });
        } else if (j > 2 && j < numberPage) {
          await page.waitForTimeout(10000);
          await page.evaluate(
            (jdxParams) => {
              document
                .querySelector("#tranFrame")
                .contentWindow.document.body.querySelector(
                  `#currAcc > a:nth-child(${jdxParams.jdx})`
                )
                .click();
              return null;
            },
            { jdx: j }
          );
        } else {
          await page.waitForTimeout(10000);
          await page.evaluate(() => {
            document
              .querySelector("#tranFrame")
              .contentWindow.document.body.querySelector(
                "#currAcc > a:nth-child(7)"
              )
              .click();
            return null;
          });
        }
        await page.waitForTimeout(3000);
        const rowCount = await page.evaluate(() => {
          const table = document
            .querySelector("#tranFrame")
            .contentWindow.document.body.querySelector("#data_content_onday")
            .querySelectorAll("tr");
          return table.length;
        });

        for (let i = 1; i < rowCount; i++) {
          const tradingCode = await page.evaluate(
            (idxParams) => {
              return document
                .querySelector("#tranFrame")
                .contentWindow.document.body.querySelector(
                  `#allResultTableBody > tr:nth-child(${idxParams.idx}) > td:nth-child(2)`
                )
                .innerHTML.trim();
            },
            { idx: i }
          );

          const declaration = await page.evaluate(
            (idxParams) => {
              return document
                .querySelector("#tranFrame")
                .contentWindow.document.body.querySelector(
                  `#allResultTableBody > tr:nth-child(${idxParams.idx}) > td:nth-child(3) > a`
                ) !== null
                ? document
                    .querySelector("#tranFrame")
                    .contentWindow.document.body.querySelector(
                      `#allResultTableBody > tr:nth-child(${idxParams.idx}) > td:nth-child(3) > a`
                    )
                    .innerHTML.trim()
                : document
                    .querySelector("#tranFrame")
                    .contentWindow.document.body.querySelector(
                      `#allResultTableBody > tr:nth-child(${idxParams.idx}) > td:nth-child(3)`
                    )
                    .innerHTML.trim();
            },
            { idx: i }
          );

          const taxPeriod = await page.evaluate(
            (idxParams) => {
              return document
                .querySelector("#tranFrame")
                .contentWindow.document.body.querySelector(
                  `#allResultTableBody > tr:nth-child(${idxParams.idx}) > td:nth-child(4)`
                )
                .innerHTML.trim();
            },
            { idx: i }
          );

          const typeOfDeclaration = await page.evaluate(
            (idxParams) => {
              return document
                .querySelector("#tranFrame")
                .contentWindow.document.body.querySelector(
                  `#allResultTableBody > tr:nth-child(${idxParams.idx}) > td:nth-child(5)`
                )
                .innerHTML.trim();
            },
            { idx: i }
          );

          const numberOfSubmissions = await page.evaluate(
            (idxParams) => {
              return document
                .querySelector("#tranFrame")
                .contentWindow.document.body.querySelector(
                  `#allResultTableBody > tr:nth-child(${idxParams.idx}) > td:nth-child(6)`
                )
                .innerHTML.trim();
            },
            { idx: i }
          );

          const numberOfAdditions = await page.evaluate(
            (idxParams) => {
              return document
                .querySelector("#tranFrame")
                .contentWindow.document.body.querySelector(
                  `#allResultTableBody > tr:nth-child(${idxParams.idx}) > td:nth-child(7)`
                )
                .innerHTML.trim();
            },
            { idx: i }
          );

          const dateOfApplication = await page.evaluate(
            (idxParams) => {
              return document
                .querySelector("#tranFrame")
                .contentWindow.document.body.querySelector(
                  `#allResultTableBody > tr:nth-child(${idxParams.idx}) > td:nth-child(8)`
                )
                .innerHTML.trim();
            },
            { idx: i }
          );

          const placeOfSubmission = await page.evaluate(
            (idxParams) => {
              return document
                .querySelector("#tranFrame")
                .contentWindow.document.body.querySelector(
                  `#allResultTableBody > tr:nth-child(${idxParams.idx}) > td:nth-child(10)`
                )
                .innerHTML.trim();
            },
            { idx: i }
          );

          const status = await page.evaluate(
            (idxParams) => {
              return document
                .querySelector("#tranFrame")
                .contentWindow.document.body.querySelector(
                  `#allResultTableBody > tr:nth-child(${idxParams.idx}) > td:nth-child(11)`
                )
                .innerHTML.trim();
            },
            { idx: i }
          );

          arrayA.push({
            tradingCode: tradingCode,
            declaration: declaration,
            taxPeriod: taxPeriod,
            typeOfDeclaration: typeOfDeclaration,
            numberOfSubmissions: numberOfSubmissions,
            numberOfAdditions: numberOfAdditions,
            dateOfApplication: dateOfApplication,
            placeOfSubmission: placeOfSubmission,
            status: status,
          });

          const downloadFile = await page.evaluate(
            (idxParams) => {
              return document
                .querySelector("#tranFrame")
                .contentWindow.document.body.querySelector(
                  `#allResultTableBody > tr:nth-child(${idxParams.idx}) > td:nth-child(3) > a`
                ) !== null
                ? document
                    .querySelector("#tranFrame")
                    .contentWindow.document.body.querySelector(
                      `#allResultTableBody > tr:nth-child(${idxParams.idx}) > td:nth-child(3) > a`
                    )
                    .click()
                : null;
            },
            { idx: i }
          );
          await page.waitForTimeout(3000);
        }
      }
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Sheet1");
      worksheet.columns = [
        { header: "Mã giao dịch", key: "tradingCode", width: 20 },
        { header: "Tờ khai/Phụ lục", key: "declaration", width: 20 },
        { header: "Kỳ tính thuế", key: "taxPeriod", width: 20 },
        { header: "Loại tờ khai", key: "typeOfDeclaration", width: 20 },
        { header: "Lần nộp", key: "numberOfSubmissions", width: 20 },
        { header: "Lần bổ sung", key: "numberOfAdditions", width: 20 },
        { header: "Ngày nộp", key: "dateOfApplication", width: 20 },
        { header: "Nơi nộp", key: "placeOfSubmission", width: 20 },
        { header: "Trạng thái", key: "status", width: 20 },
      ];
      arrayA.forEach((row) => {
        worksheet.addRow(row);
      });
      workbook.xlsx.writeFile("data.xlsx").then(() => {
        console.log("Đã xuất tệp Excel thành công");
      });
      resolve();
    } catch (error) {
      console.log("Lỗi ở scrape category: " + error);
      reject(error);
    }
  });

module.exports = {
  scrapeCategory,
};
