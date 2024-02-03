const http = require("https");
const fs = require("fs");
const xlsx = require("xlsx");

const options = {
  method: "POST",
  hostname: "api.tickertape.in",
  port: null,
  path: "/mf-screener/query",
  headers: {
    Accept: "*/*",
    "User-Agent": "Thunder Client (https://www.thunderclient.com)",
    "Content-Type": "application/json",
  },
};

const req = http.request(options, function (res) {
  const chunks = [];

  res.on("data", function (chunk) {
    chunks.push(chunk);
  });

  res.on("end", function () {
    const body = Buffer.concat(chunks);
    const result = JSON.parse(body.toString());
    let sortedRows;

    if (result.success) {
      sortedRows = result.data.result.sort((a, b) =>
        a.name.localeCompare(b.name)
      );

      const rows = sortedRows.map((fund) => {
        const values = {};
        fund.values.forEach((item) => {
          values[item.filter] = item.doubleVal || item.strVal;
        });

        return {
          name: fund.name,
          sector: fund.sector,
          ...values,
        };
      });

      // Create a worksheet
      const ws = xlsx.utils.json_to_sheet(rows);

      // Create a workbook and add the worksheet
      const wb = xlsx.utils.book_new();
      xlsx.utils.book_append_sheet(wb, ws, "Mutual Funds");

      // Save the workbook to a file
      var date = new Date().getTime();
      xlsx.writeFile(wb, `mutual_funds_${date}.xlsx`);

      console.log("Excel sheet saved successfully.");
    } else {
      console.error("API request failed");
    }
  });
});

req.write(
  JSON.stringify({
    match: { option: ["Growth", "Bonus"] },
    sortBy: "subsector",
    sortOrder: 1,
    project: [
        "subsector",
        "option",
        "aum",
        "ret1y",
        "ret3y",
        "ret5y",
        "expRatio", 
        "exitLoad",
        "ageInMon",
        "stdDevAnn",
        "riskClassification"
      ],
    offset: 0,
    count: 2000,
    mfIds: [],
  })
);

req.end();
