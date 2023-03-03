const ExcelJS = require("exceljs");
const scrapeCategory = async (browser, url) =>
  new Promise(async (resolve, reject) => {
    try {
      let taxID;
      const arrayA = [];
      let page = await browser.newPage();
      console.log(">> Mở tab mới...");

      await page.goto(url);
      await page.click(
        "#bodyP > div.khungtong > div.frm_login > div.khungbaolongin > div.left-conten > div > div:nth-child(2) > a"
      );
      await page.waitForNavigation();
      await page.click(
        "#bodyP > div > div.banner > div.dangnhap > span:nth-child(2) > button"
      );

      await page.waitForSelecter("#_userName");
      await page.focus("#_userName");
      await page.type("#_userName", "0315375348-QL");

      await page.waitForSelecter("#password");
      await page.focus("#password");
      await page.type("#password", "@Gonext$12");

      await page.waitForNavigation();
      const currentUrl = page.url();

      // const workbook = new ExcelJS.Workbook();
      // const worksheet = workbook.addWorksheet("Sheet1");
      // worksheet.columns = [
      //   { header: "Tên công ty", key: "companyName", width: 30 },
      //   { header: "Tên quốc tế", key: "companyNameEn", width: 20 },
      //   { header: "Tên viết tắt", key: "companyNameShort", width: 10 },
      //   { header: "Mã số thuế", key: "taxID", width: 30 },
      //   { header: "Địa chỉ", key: "address", width: 20 },
      //   { header: "Người đại diện", key: "owners", width: 10 },
      //   { header: "Số Điện thoại", key: "phone", width: 30 },
      //   { header: "Ngày hoạt động", key: "startDay", width: 30 },
      //   { header: "Quản lý bởi", key: "ownerBy", width: 20 },
      //   { header: "Loại hình DN", key: "legacyType", width: 10 },
      //   { header: "Tình trạng", key: "status", width: 30 },
      // ];
      // arrayA.forEach((row) => {
      //   worksheet.addRow(row);
      // });
      // workbook.xlsx.writeFile("data.xlsx").then(() => {
      //   console.log("Đã xuất tệp Excel thành công");
      // });
      resolve();
    } catch (error) {
      console.log("Lỗi ở scrape category: " + error);
      reject(error);
    }
  });

module.exports = {
  scrapeCategory,
};
