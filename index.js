const readXlsxFile = require("read-excel-file/node");
const writeXlsxFile = require("write-excel-file/node");
const fs = require("fs");

let users = {};
let items = [];

readXlsxFile("./제품내역조회.xlsx").then((rows) => {
  let usedCount = [];
  let userCount = [];
  for (let i = 2; i < rows.length; i++) {
    const inputData = {
      model: rows[i][1],
      user: rows[i][9],
    };

    if (usedCount.find((v) => v.model === inputData.model)) {
      usedCount[usedCount.findIndex((v) => v.model === inputData.model)].count++;
    } else {
      usedCount.push({ model: inputData.model, count: 1 });
    }

    if (items.find((v) => v.model !== inputData.model)) {
      items.push(inputData.model);
    }

    if (users.hasOwnProperty(inputData.user)) {
      if (users[inputData.user].items.find((v) => v.model === inputData.model)) {
        users[inputData.user].items[users[inputData.user].items.findIndex((v) => v.model === inputData.model)].count++;
      } else {
        users[inputData.user].items.push({ user: inputData.user, model: inputData.model, count: 1 });
      }
    } else {
      users[inputData.user] = { items: [{ user: inputData.user, model: inputData.model, count: 1 }] };
    }

    //   jsonData.push(inputData);
  }

  // console.log(usedCount.sort((a, b) => b.count - a.count));
  let seoul = [];
  let daejeon = [];
  let busan = [];
  let agency = [];
  let etc = [];
  for (let i in users) {
    users[i].items.sort((a, b) => b.count - a.count);
    if (
      i === "이현석" ||
      i === "강동균" ||
      i === "김선경" ||
      i === "김근욱" ||
      i === "김석구" ||
      i === "김진영" ||
      i === "박현기" ||
      i === "송성진" ||
      i === "엄근식" ||
      i === "여동구" ||
      i === "우창수" ||
      i === "정미화" ||
      i === "최현경"
    ) {
      seoul.push(users[i].items);
    } else if (i === "김성은" || i === "김효섭" || i === "안병찬") {
      daejeon.push(users[i].items);
    } else if (i === "김상훈" || i === "최지태" || i === "전해지" || i === "황준식") {
      busan.push(users[i].items);
    } else {
      agency.push(users[i].items);
    }
  }

  console.log(seoul);

  let seoulCnt = [];
  let daejeonCnt = [];
  let busanCnt = [];
  let agencyCnt = [];
  seoul.forEach((v) => {
    v.forEach((a) => {
      if (seoulCnt.find((val) => val.model === a.model)) {
        seoulCnt[seoulCnt.findIndex((val) => val.model === a.model)].count += a.count;
      } else {
        seoulCnt.push({ model: a.model, count: 1 });
      }
    });
  });
  daejeon.forEach((v) => {
    v.forEach((a) => {
      if (daejeonCnt.find((val) => val.model === a.model)) {
        daejeonCnt[daejeonCnt.findIndex((val) => val.model === a.model)].count += a.count;
      } else {
        daejeonCnt.push({ model: a.model, count: 1 });
      }
    });
  });
  busan.forEach((v) => {
    v.forEach((a) => {
      if (busanCnt.find((val) => val.model === a.model)) {
        busanCnt[busanCnt.findIndex((val) => val.model === a.model)].count += a.count;
      } else {
        busanCnt.push({ model: a.model, count: 1 });
      }
    });
  });
  agency.forEach((v) => {
    v.forEach((a) => {
      if (agencyCnt.find((val) => val.model === a.model)) {
        agencyCnt[agencyCnt.findIndex((val) => val.model === a.model)].count += a.count;
      } else {
        agencyCnt.push({ model: a.model, count: 1 });
      }
    });
  });

  seoulCnt.sort((a, b) => b.count - a.count);
  daejeonCnt.sort((a, b) => b.count - a.count);
  busanCnt.sort((a, b) => b.count - a.count);
  agencyCnt.sort((a, b) => b.count - a.count);

  let seoulData = JSON.stringify(seoulCnt, null, 2);
  let daejeonData = JSON.stringify(daejeonCnt, null, 2);
  let busanData = JSON.stringify(busanCnt, null, 2);
  let agencyData = JSON.stringify(agencyCnt, null, 2);
  fs.writeFile("seoulCountData.txt", seoulData, function (err) {
    if (err) {
      console.log(err);
    }
  });
  fs.writeFile("daejeonCountData.txt", daejeonData, function (err) {
    if (err) {
      console.log(err);
    }
  });
  fs.writeFile("busanCountData.txt", busanData, function (err) {
    if (err) {
      console.log(err);
    }
  });
  fs.writeFile("agencyCountData.txt", agencyData, function (err) {
    if (err) {
      console.log(err);
    }
  });
});
