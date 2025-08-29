/// <reference path="office-js.d.ts" />

// document.getElementById("button_a").addEventListener("click", () => tryCatch(apply));
// document.getElementById("button_b").addEventListener("click", () => tryCatch(bring));
// document.getElementById("button_c").addEventListener("click", () => tryCatch(combine));
// document.getElementById("button_d").addEventListener("click", () => tryCatch(deliver));
// document.getElementById("button_e").addEventListener("click", () => tryCatch(explore));
// document.getElementById("button_f").addEventListener("click", () => tryCatch(fix));

const button_a = document.getElementById("button_a")!;
button_a.addEventListener("click", () => tryCatch(apply));

const button_b = document.getElementById("button_b")!;
button_b.addEventListener("click", () => tryCatch(bring));

const button_c = document.getElementById("button_c")!;
button_c.addEventListener("click", () => tryCatch(combine));

const button_d = document.getElementById("button_d")!;
button_d.addEventListener("click", () => tryCatch(deliver));

const button_e = document.getElementById("button_e")!;
button_e.addEventListener("click", () => tryCatch(explore));

const button_f = document.getElementById("button_f")!;
button_f.addEventListener("click", () => tryCatch(fix));

// 데이터 반영
async function apply() {
  await Excel.run(async (context) => {
    const applysheet = context.workbook.worksheets.getItem("반영");
    const applysheetrange = applysheet.getRange("C6:K20");
    applysheetrange.load("values");

    const settingsheet = context.workbook.worksheets.getItem("설정");
    const usersettingrange = settingsheet.getRange("I13:I16");
    usersettingrange.load("values");

    const datasheet = context.workbook.worksheets.getItem("데이터");
    const datasheetindexrange = datasheet.getRange("A:C").getUsedRange();
    datasheetindexrange.load("values");

    const logsheet = context.workbook.worksheets.getItem("로그");
    const logsheetrange = logsheet.getRange("A:R").getUsedRange();
    logsheetrange.load("values");

    await context.sync();

    const username = usersettingrange.values[0][0] as string;
    const applysheetdata = applysheetrange.values;
    const logsheetdata = logsheetrange.values;

    // 새 데이터 작성:
    let newdata: (string | number | boolean)[] = [];

    // 새 로그 작성:
    let newlog: (string | number | boolean)[] = [];
    // newlog 구조:
    // [0]날짜와 시간, [1]사용자, [2]변경 유형, [3]도착지, [4]운송사, [5]운송 단위, [6]변경한 항목, [7~11]변경 전 값, [12~16]변경 후 값, [17](선택)
    //
    // 변경한 항목의 카테고리:
    // 유효 기간, 운임 특성, 부산발 정보, 인천발 정보, 광양발 정보, 평택-당진발 정보, 경로 정보, 메모, 프리타임, 도착지 비용1~7 (총 16개)

    // 도착 위치, 운송사, 운송 단위
    newdata = newdata.concat(applysheetdata[0][0], applysheetdata[0][1]);
    newdata = newdata.concat((applysheetdata[0][2] as boolean) ? "LCL" : "FCL");
    const newdataindex = newdata.map(String).join("^"); //데이터 시트에 입력할 행을 찾을 때 사용할 참조값

    // 유효 기간 부터, 까지, 운임 특성
    let actualdate = new Date();
    if ((applysheetdata[3][0] as string) !== "") {
      let numerifieddate = applysheetdata[3][0];
      numerifieddate -= 25569;
      numerifieddate *= 86400000;
      actualdate = new Date(numerifieddate);
    }
    let stringifieddate =
      actualdate.getFullYear() +
      "-" +
      ("0" + (actualdate.getMonth() + 1)).slice(-2) +
      "-" +
      ("0" + actualdate.getDate()).slice(-2);
    newdata = newdata.concat(stringifieddate);

    actualdate = new Date();
    let targetyear = actualdate.getFullYear() as number;
    let targetmonth = actualdate.getMonth() as number;
    let targetdate = actualdate.getDate() as number;
    if (targetdate >= 15) {
      if (targetmonth === 11) {
        targetmonth = 0;
        targetyear++;
      } else {
        targetmonth++;
      }
      targetdate = 14;
    } else {
      if (targetmonth === 1) {
        if ((targetyear % 4 === 0 && targetyear % 100 !== 0) || targetyear % 400 === 0) {
          targetdate = 29;
        } else {
          targetdate = 28;
        }
      } else {
        if ([1, 3, 5, 7, 8, 10, 12].indexOf(targetmonth as number) !== -1) {
          targetdate = 31;
        } else {
          targetdate = 30;
        }
      }
    }
    if ((applysheetdata[3][1] as string) !== "") {
      let numerifieddate = Math.round(applysheetdata[3][1]);
      numerifieddate -= 25569;
      numerifieddate *= 86400000;
      actualdate = new Date(numerifieddate);
    } else {
      actualdate = new Date(targetyear, targetmonth, targetdate);
    }
    stringifieddate =
      actualdate.getFullYear() +
      "-" +
      ("0" + (actualdate.getMonth() + 1)).slice(-2) +
      "-" +
      ("0" + actualdate.getDate()).slice(-2);
    newdata = newdata.concat(stringifieddate, applysheetdata[3][2]);

    // 부산발, 인천발, 광양발, 평택-당진발 운임 및 소요일 정보
    newdata = newdata.concat(applysheetdata[6][0] as string, applysheetdata[6][1] as string, applysheetdata[6][2] as string);
    newdata = newdata.concat(applysheetdata[7][0] as string, applysheetdata[7][1] as string, applysheetdata[7][2] as string);
    newdata = newdata.concat(applysheetdata[8][0] as string, applysheetdata[8][1] as string, applysheetdata[8][2] as string);
    newdata = newdata.concat(applysheetdata[9][0] as string, applysheetdata[9][1] as string, applysheetdata[9][2] as string);

    // 경로 정보
    newdata = newdata.concat(applysheetdata[12][0], applysheetdata[12][1], applysheetdata[12][2]);

    // 도착 위치 프리타임
    newdata = newdata.concat(applysheetdata[3][4], applysheetdata[3][5]);

    actualdate = new Date("9999-12-31");
    if ((applysheetdata[3][6] as string) !== "") {
      let numerifieddate = Math.round(applysheetdata[3][6]);
      numerifieddate -= 25569;
      numerifieddate *= 86400000;
      actualdate = new Date(numerifieddate);
    }
    stringifieddate =
      actualdate.getFullYear() +
      "-" +
      ("0" + (actualdate.getMonth() + 1)).slice(-2) +
      "-" +
      ("0" + actualdate.getDate()).slice(-2);
    newdata = newdata.concat(stringifieddate);

    // 도착지 비용 1~7
    newdata = newdata.concat(
      applysheetdata[6][4],
      applysheetdata[6][5],
      applysheetdata[6][6],
      applysheetdata[6][7],
      applysheetdata[6][8],
    );
    newdata = newdata.concat(
      applysheetdata[7][4],
      applysheetdata[7][5],
      applysheetdata[7][6],
      applysheetdata[7][7],
      applysheetdata[7][8],
    );
    newdata = newdata.concat(
      applysheetdata[8][4],
      applysheetdata[8][5],
      applysheetdata[8][6],
      applysheetdata[8][7],
      applysheetdata[8][8],
    );
    newdata = newdata.concat(
      applysheetdata[9][4],
      applysheetdata[9][5],
      applysheetdata[9][6],
      applysheetdata[9][7],
      applysheetdata[9][8],
    );
    newdata = newdata.concat(
      applysheetdata[10][4],
      applysheetdata[10][5],
      applysheetdata[10][6],
      applysheetdata[10][7],
      applysheetdata[10][8],
    );
    newdata = newdata.concat(
      applysheetdata[11][4],
      applysheetdata[11][5],
      applysheetdata[11][6],
      applysheetdata[11][7],
      applysheetdata[11][8],
    );
    newdata = newdata.concat(
      applysheetdata[12][4],
      applysheetdata[12][5],
      applysheetdata[12][6],
      applysheetdata[12][7],
      applysheetdata[12][8],
    );

    // 메모
    newdata = newdata.concat(applysheetdata[14][0]);

    // 참조값으로 입력할 행 찾기
    const datasheetindexdata = datasheetindexrange.values.map((row) => row.slice(0, 3).map(String).join("^"));
    let datasheetinputindex = datasheetindexdata.indexOf(newdataindex, 2);
    if (datasheetinputindex === -1 || datasheetindexdata.length === 2) {
      datasheetinputindex = datasheetindexdata.length;
    }
    datasheetinputindex++;

    // 입력
    const datasheetinputrange = datasheet.getRange("A" + datasheetinputindex + ":BH" + datasheetinputindex);
    datasheetinputrange.load("values");
    await context.sync();

    newlog[0] = new Date().toISOString();
    newlog[1] = username;
    newlog[3] = newdata[0];
    newlog[4] = newdata[1];
    newlog[5] = newdata[2];
    newlog[17] = false;
    if (datasheetinputrange.values[0][0] === "") {
      newlog[2] = "신규";
      newlog[6] = "전체";
      newlog[7] = "";
      newlog[8] = "";
      newlog[9] = "";
      newlog[10] = "";
      newlog[11] = "";
      newlog[12] = "";
      newlog[13] = "";
      newlog[14] = "";
      newlog[15] = "";
      newlog[16] = "";
      logsheetdata.push([...newlog]);
    } else {
      newlog[2] = "수정";
      if (datasheetinputrange.values[0][3] !== newdata[3] || datasheetinputrange.values[0][4] !== newdata[4]) {
        newlog[6] = "유효 기간";
        newlog[7] = String(datasheetinputrange.values[0][3]);
        newlog[8] = String(datasheetinputrange.values[0][4]);
        newlog[9] = "";
        newlog[10] = "";
        newlog[11] = "";
        newlog[12] = String(newdata[3]);
        newlog[13] = String(newdata[4]);
        newlog[14] = "";
        newlog[15] = "";
        newlog[16] = "";
        logsheetdata.push([...newlog]);
      }
      if (datasheetinputrange.values[0][5] !== newdata[5]) {
        newlog[6] = "운임 특성";
        newlog[7] = String(datasheetinputrange.values[0][5]);
        newlog[8] = "";
        newlog[9] = "";
        newlog[10] = "";
        newlog[11] = "";
        newlog[12] = String(newdata[5]);
        newlog[13] = "";
        newlog[14] = "";
        newlog[15] = "";
        newlog[16] = "";
        logsheetdata.push([...newlog]);
      }
      if (
        datasheetinputrange.values[0][6] !== newdata[6] || 
        datasheetinputrange.values[0][7] !== newdata[7] || 
        datasheetinputrange.values[0][8] !== newdata[8]
      ) {
        newlog[6] = "부산발 정보";
        newlog[7] = String(datasheetinputrange.values[0][6]);
        newlog[8] = String(datasheetinputrange.values[0][7]);
        newlog[9] = String(datasheetinputrange.values[0][8]);
        newlog[10] = "";
        newlog[11] = "";
        newlog[12] = String(newdata[6]);
        newlog[13] = String(newdata[7]);
        newlog[14] = String(newdata[8]);
        newlog[15] = "";
        newlog[16] = "";
        logsheetdata.push([...newlog]);
      }
      if (
        datasheetinputrange.values[0][9] !== newdata[9] || 
        datasheetinputrange.values[0][10] !== newdata[10] || 
        datasheetinputrange.values[0][11] !== newdata[11]
      ) {
        newlog[6] = "인천발 정보";
        newlog[7] = String(datasheetinputrange.values[0][9]);
        newlog[8] = String(datasheetinputrange.values[0][10]);
        newlog[9] = String(datasheetinputrange.values[0][11]);
        newlog[10] = "";
        newlog[11] = "";
        newlog[12] = String(newdata[9]);
        newlog[13] = String(newdata[10]);
        newlog[14] = String(newdata[11]);
        newlog[15] = "";
        newlog[16] = "";
        logsheetdata.push([...newlog]);
      }
      if (
        datasheetinputrange.values[0][12] !== newdata[12] || 
        datasheetinputrange.values[0][13] !== newdata[13] || 
        datasheetinputrange.values[0][14] !== newdata[14]
      ) {
        newlog[6] = "광양발 정보";
        newlog[7] = String(datasheetinputrange.values[0][12]);
        newlog[8] = String(datasheetinputrange.values[0][13]);
        newlog[9] = String(datasheetinputrange.values[0][14]);
        newlog[10] = "";
        newlog[11] = "";
        newlog[12] = String(newdata[12]);
        newlog[13] = String(newdata[13]);
        newlog[14] = String(newdata[14]);
        newlog[15] = "";
        newlog[16] = "";
        logsheetdata.push([...newlog]);
      }
      if (
        datasheetinputrange.values[0][15] !== newdata[15] ||
        datasheetinputrange.values[0][16] !== newdata[16] ||
        datasheetinputrange.values[0][17] !== newdata[17]
      ) {
        newlog[6] = "평택-당진발 정보";
        newlog[7] = String(datasheetinputrange.values[0][15]);
        newlog[8] = String(datasheetinputrange.values[0][16]);
        newlog[9] = String(datasheetinputrange.values[0][17]);
        newlog[10] = "";
        newlog[11] = "";
        newlog[12] = String(newdata[15]);
        newlog[13] = String(newdata[16]);
        newlog[14] = String(newdata[17]);
        newlog[15] = "";
        newlog[16] = "";
        logsheetdata.push([...newlog]);
      }
      if (
        datasheetinputrange.values[0][18] !== newdata[18] ||
        datasheetinputrange.values[0][19] !== newdata[19] ||
        datasheetinputrange.values[0][20] !== newdata[20]
      ) {
        newlog[6] = "경로 정보";
        newlog[7] = String(datasheetinputrange.values[0][18]);
        newlog[8] = String(datasheetinputrange.values[0][19]);
        newlog[9] = String(datasheetinputrange.values[0][20]);
        newlog[10] = "";
        newlog[11] = "";
        newlog[12] = String(newdata[18]);
        newlog[13] = String(newdata[19]);
        newlog[14] = String(newdata[20]);
        newlog[15] = "";
        newlog[16] = "";
        logsheetdata.push([...newlog]);
      }
      if (
        datasheetinputrange.values[0][21] !== newdata[21] ||
        datasheetinputrange.values[0][22] !== newdata[22] ||
        datasheetinputrange.values[0][23] !== newdata[23]
      ) {
        newlog[6] = "프리타임";
        newlog[7] = String(datasheetinputrange.values[0][21]);
        newlog[8] = String(datasheetinputrange.values[0][22]);
        newlog[9] = String(datasheetinputrange.values[0][23]);
        newlog[10] = "";
        newlog[11] = "";
        newlog[12] = String(newdata[21]);
        newlog[13] = String(newdata[22]);
        newlog[14] = String(newdata[23]);
        newlog[15] = "";
        newlog[16] = "";
        logsheetdata.push([...newlog]);
      }
      for (let i = 1; i <= 7; i++) {
        if (
          datasheetinputrange.values[0][24 + ((i - 1) * 5)] !== newdata[24 + ((i - 1) * 5)] ||
          datasheetinputrange.values[0][25 + ((i - 1) * 5)] !== newdata[25 + ((i - 1) * 5)] ||
          datasheetinputrange.values[0][26 + ((i - 1) * 5)] !== newdata[26 + ((i - 1) * 5)] ||
          datasheetinputrange.values[0][27 + ((i - 1) * 5)] !== newdata[27 + ((i - 1) * 5)] ||
          datasheetinputrange.values[0][28 + ((i - 1) * 5)] !== newdata[28 + ((i - 1) * 5)]
        ) {
          newlog[6] = "도착지 비용" + i;
          newlog[7] = String(datasheetinputrange.values[0][24 + ((i - 1) * 5)]);
          newlog[8] = String(datasheetinputrange.values[0][25 + ((i - 1) * 5)]);
          newlog[9] = String(datasheetinputrange.values[0][26 + ((i - 1) * 5)]);
          newlog[10] = String(datasheetinputrange.values[0][27 + ((i - 1) * 5)]);
          newlog[11] = String(datasheetinputrange.values[0][28 + ((i - 1) * 5)]);
          newlog[12] = String(newdata[24 + ((i - 1) * 5)]);
          newlog[13] = String(newdata[25 + ((i - 1) * 5)]);
          newlog[14] = String(newdata[26 + ((i - 1) * 5)]);
          newlog[15] = String(newdata[27 + ((i - 1) * 5)]);
          newlog[16] = String(newdata[28 + ((i - 1) * 5)]);
          logsheetdata.push([...newlog]);
        }
      }
      if (datasheetinputrange.values[0][59] !== newdata[59]) {
        newlog[6] = "메모";
        newlog[7] = String(datasheetinputrange.values[0][59]);
        newlog[8] = "";
        newlog[9] = "";
        newlog[10] = "";
        newlog[11] = "";
        newlog[12] = String(newdata[59]);
        newlog[13] = "";
        newlog[14] = "";
        newlog[15] = "";
        newlog[16] = "";
        logsheetdata.push([...newlog]);
      }
    }

    datasheetinputrange.values = [newdata];

    const logsheetnewrange = context.workbook.worksheets.getItem("로그").getRange("A1:R" + logsheetdata.length);
    logsheetnewrange.load("values");
    logsheetnewrange.getColumn(17).control = {type: Excel.CellControlType.checkbox};
    await context.sync();

    logsheetnewrange.values = logsheetdata;
  });
}

// 데이터 참조
async function bring() {
  await Excel.run(async (context) => {
    const applysheet = context.workbook.worksheets.getItem("반영");
    const applysheetrange = applysheet.getRange("C6:K20");
    const applysheetformularange1 = applysheet.getRange("G6:J6");
    const applysheetformularange2 = applysheet.getRange("C11:D11");
    applysheetrange.load("values");
    applysheetformularange1.load("formulas");
    applysheetformularange2.load("formulas");

    const datasheet = context.workbook.worksheets.getItem("데이터");
    const datasheetindexrange = datasheet.getRange("A:C").getUsedRange();
    datasheetindexrange.load("values");

    await context.sync();

    const applysheetdata = applysheetrange.values;
    const applysheetformula1 = applysheetformularange1.formulas;
    const applysheetformula2 = applysheetformularange2.formulas;

    const basicdata = [applysheetdata[0][0], applysheetdata[0][1]].concat(
      applysheetdata[0][2] as boolean ? "LCL" : "FCL",
    );
    const basicdataindex = basicdata.map(String).join("^");

    const datasheetindexdata = datasheetindexrange.values.map((row) => row.slice(0, 3).map(String).join("^"));
    let datasheetindex = datasheetindexdata.indexOf(basicdataindex, 2);
    if (datasheetindex === -1 || datasheetindexdata.length === 2) {
      datasheetindex = datasheetindexdata.length;
    }
    datasheetindex++;

    const datasheetrange = datasheet.getRange("A" + datasheetindex + ":BH" + datasheetindex);
    datasheetrange.load("values");
    await context.sync();
    const datasheetdata = (datasheetrange.values)[0];

    if (datasheetdata[0] as string === "") {
      applysheetdata[3][2] = "신규";
    } else {
      // 도착 위치, 운송사, 운송 단위
      /** 필요 없음
      applysheetdata[0][0] = datasheetdata[0] as string;
      applysheetdata[0][1] = datasheetdata[1] as string;
      if (datasheetdata[2] as string === "LCL") {
        applysheetdata[0][2] = true;
      } else {
        applysheetdata[0][2] = false;
      }
      */
      // 유효 기간 부터, 까지, 운임 특성
      applysheetdata[3][0] = datasheetdata[3] as string;
      applysheetdata[3][1] = datasheetdata[4] as string;
      applysheetdata[3][2] = datasheetdata[5] as string;
      // 부산발, 인천발, 광양발, 평택-당진발 운임 및 소요일 정보
      applysheetdata[6][0] = datasheetdata[6] as string;
      applysheetdata[6][1] = datasheetdata[7] as string;
      applysheetdata[6][2] = datasheetdata[8] as string;
      applysheetdata[7][0] = datasheetdata[9] as string;
      applysheetdata[7][1] = datasheetdata[10] as string;
      applysheetdata[7][2] = datasheetdata[11] as string;
      applysheetdata[8][0] = datasheetdata[12] as string;
      applysheetdata[8][1] = datasheetdata[13] as string;
      applysheetdata[8][2] = datasheetdata[14] as string;
      applysheetdata[9][0] = datasheetdata[15] as string;
      applysheetdata[9][1] = datasheetdata[16] as string;
      applysheetdata[9][2] = datasheetdata[17] as string;
      // 경로 정보
      applysheetdata[12][0] = datasheetdata[18] as string;
      applysheetdata[12][1] = datasheetdata[19] as string;
      applysheetdata[12][2] = datasheetdata[20] as string;
      // 도착 위치 프리타임
      applysheetdata[3][4] = datasheetdata[21] as string;
      applysheetdata[3][5] = datasheetdata[22] as string;
      applysheetdata[3][6] = datasheetdata[23] as string;
      // 도착지 비용 1~7
      applysheetdata[6][4] = datasheetdata[24] as boolean;
      applysheetdata[6][5] = datasheetdata[25] as string;
      applysheetdata[6][6] = datasheetdata[26] as string;
      applysheetdata[6][7] = datasheetdata[27] as string;
      applysheetdata[6][8] = datasheetdata[28] as number;
      applysheetdata[7][4] = datasheetdata[29] as boolean;
      applysheetdata[7][5] = datasheetdata[30] as string;
      applysheetdata[7][6] = datasheetdata[31] as string;
      applysheetdata[7][7] = datasheetdata[32] as string;
      applysheetdata[7][8] = datasheetdata[33] as number;
      applysheetdata[8][4] = datasheetdata[34] as boolean;
      applysheetdata[8][5] = datasheetdata[35] as string;
      applysheetdata[8][6] = datasheetdata[36] as string;
      applysheetdata[8][7] = datasheetdata[37] as string;
      applysheetdata[8][8] = datasheetdata[38] as number;
      applysheetdata[9][4] = datasheetdata[39] as boolean;
      applysheetdata[9][5] = datasheetdata[40] as string;
      applysheetdata[9][6] = datasheetdata[41] as string;
      applysheetdata[9][7] = datasheetdata[42] as string;
      applysheetdata[9][8] = datasheetdata[43] as number;
      applysheetdata[10][4] = datasheetdata[44] as boolean;
      applysheetdata[10][5] = datasheetdata[45] as string;
      applysheetdata[10][6] = datasheetdata[46] as string;
      applysheetdata[10][7] = datasheetdata[47] as string;
      applysheetdata[10][8] = datasheetdata[48] as number;
      applysheetdata[11][4] = datasheetdata[49] as boolean;
      applysheetdata[11][5] = datasheetdata[50] as string;
      applysheetdata[11][6] = datasheetdata[51] as string;
      applysheetdata[11][7] = datasheetdata[52] as string;
      applysheetdata[11][8] = datasheetdata[53] as number;
      applysheetdata[12][4] = datasheetdata[54] as boolean;
      applysheetdata[12][5] = datasheetdata[55] as string;
      applysheetdata[12][6] = datasheetdata[56] as string;
      applysheetdata[12][7] = datasheetdata[57] as string;
      applysheetdata[12][8] = datasheetdata[58] as number;
      // 메모
      applysheetdata[14][0] = datasheetdata[59] as string;
    }

    applysheetrange.values = applysheetdata;

    // 수식 덮어쓰기
    applysheetformularange1.formulas = applysheetformula1;
    applysheetformularange2.formulas = applysheetformula2;
  });
}

// 데이터 병합
async function combine() {
  await Excel.run(async (context) => {
    const settingsheet = context.workbook.worksheets.getItem("설정");
    const usersettingrange = settingsheet.getRange("I13:I16");
    usersettingrange.load("values");

    const datasheet = context.workbook.worksheets.getItem("데이터");
    const datasheetrange = datasheet.getUsedRange();
    datasheetrange.load("values");

    const pusheddatasheet = context.workbook.worksheets.getItem("병합");
    const pusheddatasheetrange = pusheddatasheet.getUsedRange();
    pusheddatasheetrange.load("values");

    const logsheet = context.workbook.worksheets.getItem("로그");
    const logsheetrange = logsheet.getRange("A:R").getUsedRange();
    logsheetrange.load("values");

    await context.sync();

    const pusheddatausername = usersettingrange.values[1][0] as string;
    const pusheddataoverrides = usersettingrange.values[2][0] as boolean;
    const pusheddatasheetdata = pusheddatasheetrange.values;
    const logsheetdata = logsheetrange.values;

    let datasheetdata = datasheetrange.values;
    let singlepusheddata: (string | number | boolean)[] = [];
    let pushindexdata = "";
    let pushindex = -1;

    // 새 로그 작성:
    let newlog: (string | number | boolean)[] = [];
    // newlog 구조:
    // [0]날짜와 시간, [1]사용자, [2]변경 유형, [3]도착지, [4]운송사, [5]운송 단위, [6]변경한 항목, [7~11]변경 전 값, [12~16]변경 후 값, [17](선택)
    //
    // 변경한 항목의 카테고리:
    // 유효 기간, 운임 특성, 부산발 정보, 인천발 정보, 광양발 정보, 평택-당진발 정보, 경로 정보, 메모, 프리타임, 도착지 비용1~7 (총 16개)

    for (let i = 2; i < pusheddatasheetdata.length; i++) {
      singlepusheddata = pusheddatasheetdata[i];

      // 참조값으로 입력할 행 찾기
      pushindexdata = singlepusheddata.slice(0, 3).map(String).join("^");
      pushindex = datasheetdata.map((row) => row.slice(0, 3).map(String).join("^")).indexOf(pushindexdata, 2);
      if (pushindex === -1 || datasheetdata.length === 2) {
        pushindex = datasheetdata.length;
        datasheetdata[pushindex] = new Array(60).fill("");
      } else if (pusheddataoverrides === false) {
        continue;
      } else if (singlepusheddata === datasheetdata[pushindex]) {
        continue;
      }

      // 로그 작성
      newlog[0] = new Date().toISOString();
      newlog[1] = pusheddatausername;
      newlog[3] = singlepusheddata[0];
      newlog[4] = singlepusheddata[1];
      newlog[5] = singlepusheddata[2];
      newlog[17] = false;
      if (datasheetdata[pushindex][0] === "") {
        newlog[2] = "신규";
        newlog[6] = "전체";
        newlog[7] = "";
        newlog[8] = "";
        newlog[9] = "";
        newlog[10] = "";
        newlog[11] = "";
        newlog[12] = "";
        newlog[13] = "";
        newlog[14] = "";
        newlog[15] = "";
        newlog[16] = "";
        logsheetdata.push([...newlog]);
      } else {
        newlog[2] = "수정";
        if (datasheetdata[pushindex][3] !== singlepusheddata[3] || datasheetdata[pushindex][4] !== singlepusheddata[4]) {
          newlog[6] = "유효 기간";
          newlog[7] = String(datasheetdata[pushindex][3]);
          newlog[8] = String(datasheetdata[pushindex][4]);
          newlog[9] = "";
          newlog[10] = "";
          newlog[11] = "";
          newlog[12] = String(singlepusheddata[3]);
          newlog[13] = String(singlepusheddata[4]);
          newlog[14] = "";
          newlog[15] = "";
          newlog[16] = "";
          logsheetdata.push([...newlog]);
        }
        if (datasheetdata[pushindex][5] !== singlepusheddata[5]) {
          newlog[6] = "운임 특성";
          newlog[7] = String(datasheetdata[pushindex][5]);
          newlog[8] = "";
          newlog[9] = "";
          newlog[10] = "";
          newlog[11] = "";
          newlog[12] = String(singlepusheddata[5]);
          newlog[13] = "";
          newlog[14] = "";
          newlog[15] = "";
          newlog[16] = "";
          logsheetdata.push([...newlog]);
        }
        if (
          datasheetdata[pushindex][6] !== singlepusheddata[6] ||
          datasheetdata[pushindex][7] !== singlepusheddata[7] ||
          datasheetdata[pushindex][8] !== singlepusheddata[8]
        ) {
          newlog[6] = "부산발 정보";
          newlog[7] = String(datasheetdata[pushindex][6]);
          newlog[8] = String(datasheetdata[pushindex][7]);
          newlog[9] = String(datasheetdata[pushindex][8]);
          newlog[10] = "";
          newlog[11] = "";
          newlog[12] = String(singlepusheddata[6]);
          newlog[13] = String(singlepusheddata[7]);
          newlog[14] = String(singlepusheddata[8]);
          newlog[15] = "";
          newlog[16] = "";
          logsheetdata.push([...newlog]);
        }
        if (
          datasheetdata[pushindex][9] !== singlepusheddata[9] ||
          datasheetdata[pushindex][10] !== singlepusheddata[10] ||
          datasheetdata[pushindex][11] !== singlepusheddata[11]
        ) {
          newlog[6] = "인천발 정보";
          newlog[7] = String(datasheetdata[pushindex][9]);
          newlog[8] = String(datasheetdata[pushindex][10]);
          newlog[9] = String(datasheetdata[pushindex][11]);
          newlog[10] = "";
          newlog[11] = "";
          newlog[12] = String(singlepusheddata[9]);
          newlog[13] = String(singlepusheddata[10]);
          newlog[14] = String(singlepusheddata[11]);
          newlog[15] = "";
          newlog[16] = "";
          logsheetdata.push([...newlog]);
        }
        if (
          datasheetdata[pushindex][12] !== singlepusheddata[12] ||
          datasheetdata[pushindex][13] !== singlepusheddata[13] ||
          datasheetdata[pushindex][14] !== singlepusheddata[14]
        ) {
          newlog[6] = "광양발 정보";
          newlog[7] = String(datasheetdata[pushindex][12]);
          newlog[8] = String(datasheetdata[pushindex][13]);
          newlog[9] = String(datasheetdata[pushindex][14]);
          newlog[10] = "";
          newlog[11] = "";
          newlog[12] = String(singlepusheddata[12]);
          newlog[13] = String(singlepusheddata[13]);
          newlog[14] = String(singlepusheddata[14]);
          newlog[15] = "";
          newlog[16] = "";
          logsheetdata.push([...newlog]);
        }
        if (
          datasheetdata[pushindex][15] !== singlepusheddata[15] ||
          datasheetdata[pushindex][16] !== singlepusheddata[16] ||
          datasheetdata[pushindex][17] !== singlepusheddata[17]
        ) {
          newlog[6] = "평택-당진발 정보";
          newlog[7] = String(datasheetdata[pushindex][15]);
          newlog[8] = String(datasheetdata[pushindex][16]);
          newlog[9] = String(datasheetdata[pushindex][17]);
          newlog[10] = "";
          newlog[11] = "";
          newlog[12] = String(singlepusheddata[15]);
          newlog[13] = String(singlepusheddata[16]);
          newlog[14] = String(singlepusheddata[17]);
          newlog[15] = "";
          newlog[16] = "";
          logsheetdata.push([...newlog]);
        }
        if (
          datasheetdata[pushindex][18] !== singlepusheddata[18] ||
          datasheetdata[pushindex][19] !== singlepusheddata[19] ||
          datasheetdata[pushindex][20] !== singlepusheddata[20]
        ) {
          newlog[6] = "경로 정보";
          newlog[7] = String(datasheetdata[pushindex][18]);
          newlog[8] = String(datasheetdata[pushindex][19]);
          newlog[9] = String(datasheetdata[pushindex][20]);
          newlog[10] = "";
          newlog[11] = "";
          newlog[12] = String(singlepusheddata[18]);
          newlog[13] = String(singlepusheddata[19]);
          newlog[14] = String(singlepusheddata[20]);
          newlog[15] = "";
          newlog[16] = "";
          logsheetdata.push([...newlog]);
        }
        if (
          datasheetdata[pushindex][21] !== singlepusheddata[21] ||
          datasheetdata[pushindex][22] !== singlepusheddata[22] ||
          datasheetdata[pushindex][23] !== singlepusheddata[23]
        ) {
          newlog[6] = "프리타임";
          newlog[7] = String(datasheetdata[pushindex][21]);
          newlog[8] = String(datasheetdata[pushindex][22]);
          newlog[9] = String(datasheetdata[pushindex][23]);
          newlog[10] = "";
          newlog[11] = "";
          newlog[12] = String(singlepusheddata[21]);
          newlog[13] = String(singlepusheddata[22]);
          newlog[14] = String(singlepusheddata[23]);
          newlog[15] = "";
          newlog[16] = "";
          logsheetdata.push([...newlog]);
        }
        for (let ii = 1; ii <= 7; ii++) {
          if (
            datasheetdata[pushindex][24 + ((ii - 1) * 5)] !== singlepusheddata[24 + ((ii - 1) * 5)] ||
            datasheetdata[pushindex][25 + ((ii - 1) * 5)] !== singlepusheddata[25 + ((ii - 1) * 5)] ||
            datasheetdata[pushindex][26 + ((ii - 1) * 5)] !== singlepusheddata[26 + ((ii - 1) * 5)] ||
            datasheetdata[pushindex][27 + ((ii - 1) * 5)] !== singlepusheddata[27 + ((ii - 1) * 5)] ||
            datasheetdata[pushindex][28 + ((ii - 1) * 5)] !== singlepusheddata[28 + ((ii - 1) * 5)]
          ) {
            newlog[6] = "도착지 비용" + ii;
            newlog[7] = String(datasheetdata[pushindex][24 + ((ii - 1) * 5)]);
            newlog[8] = String(datasheetdata[pushindex][25 + ((ii - 1) * 5)]);
            newlog[9] = String(datasheetdata[pushindex][26 + ((ii - 1) * 5)]);
            newlog[10] = String(datasheetdata[pushindex][27 + ((ii - 1) * 5)]);
            newlog[11] = String(datasheetdata[pushindex][28 + ((ii - 1) * 5)]);
            newlog[12] = String(singlepusheddata[24 + ((ii - 1) * 5)]);
            newlog[13] = String(singlepusheddata[25 + ((ii - 1) * 5)]);
            newlog[14] = String(singlepusheddata[26 + ((ii - 1) * 5)]);
            newlog[15] = String(singlepusheddata[27 + ((ii - 1) * 5)]);
            newlog[16] = String(singlepusheddata[28 + ((ii - 1) * 5)]);
            logsheetdata.push([...newlog]);
          }
        }
        if (datasheetdata[pushindex][59] !== singlepusheddata[59]) {
          newlog[6] = "메모";
          newlog[7] = String(datasheetdata[pushindex][59]);
          newlog[8] = "";
          newlog[9] = "";
          newlog[10] = "";
          newlog[11] = "";
          newlog[12] = String(singlepusheddata[59]);
          newlog[13] = "";
          newlog[14] = "";
          newlog[15] = "";
          newlog[16] = "";
          logsheetdata.push([...newlog]);
        }
      }
      datasheetdata[pushindex] = singlepusheddata;
    }

    const datasheetnewrange = datasheet.getRange("A1:BH" + datasheetdata.length);
    datasheetnewrange.load("values");

    const logsheetnewrange = context.workbook.worksheets.getItem("로그").getRange("A1:R" + logsheetdata.length);
    logsheetnewrange.load("values");
    logsheetnewrange.getColumn(17).control = {type: Excel.CellControlType.checkbox};
    await context.sync();

    datasheetnewrange.values = datasheetdata;
    logsheetnewrange.values = logsheetdata;
  });
}

// 운임 시트 작성
async function deliver() {
  await Excel.run(async (context) => {
    const settingsheet = context.workbook.worksheets.getItem("설정");
    const marginsettingsrange = settingsheet.getRange("C14:D16");
    marginsettingsrange.load("values");

    const datasheet = context.workbook.worksheets.getItem("데이터");
    const datasheetrange = datasheet.getUsedRange();
    datasheetrange.load("values");

    const exchangeratesheet = context.workbook.worksheets.getItem("환율");
    const exchangeratesheetrange = exchangeratesheet.getRange("B:C").getUsedRange();
    exchangeratesheetrange.load("values");

    const deliversheet = context.workbook.worksheets.getItem("운임");
    const deliversheetheaderrange = deliversheet.getRange("A1:AT2");
    deliversheetheaderrange.load("values");

    const locationsheet = context.workbook.worksheets.getItem("위치");
    const locationsheetrange = locationsheet.getRange("A:E").getUsedRange();
    locationsheetrange.load("values");

    await context.sync();

    const marginsetting = marginsettingsrange.values;
    // marginsetting 구조: [0][0]20피트 마진율, [0][1]20피트 최소 마진, [1][0]40피트 마진율, [1][1]40피트 최소 마진,
    // [2][0]LCL 마진율, [2][1]LCL 최소 마진
    const datasheetdata = datasheetrange.values;
    const exchangeratelist = exchangeratesheetrange.values.slice(1, undefined);
    // exchangeratelist 구조: [0]환율, [1]화폐
    const deliversheetdata = deliversheetheaderrange.values;
    const locationlist = locationsheetrange.values.slice(2, undefined);
    // locationlist 구조: [0]6대륙, [1]국가, [2]지역, [3]도시/항구, [4]CITY/PORT

    let singleitem: (string | number | boolean)[] = [];
    // singleitem 구조: [0]운송사, [1]출발 위치, [2]POL, [3]도착 국가, [4]도착 위치, [5]POD,
    // [6]환적유무, [7]환적항1, [8]환적항2, [9]운송 단위, [10]운송 소요일, [11]유효기간 시작, [12]유효기간 종료,
    // [13]20피트 운임, [14]40피트 하이큐브 운임, [15]LCL 운임, [16]프리타임, [17]USCAN,
    // [18, 19, 20, 21]도착지비용1{항목, 단위, 화폐, 금액} ~ [42, 43, 44, 45]도착지비용7

    let singledata: (string | number | boolean)[] = [];
    // singledata 구조: [0]도착 위치, [1]운송사, [2]운송 단위, [3]유효기간 시작, [4]유효기간 종료, [5]운임 특성,
    // [6, 7, 8]부산발 정보{운임1, 운임2, 소요일}, [9, 10, 11]인천발 정보, [12, 13, 14]광양발 정보, [15, 16, 17]평택-당진발 정보,
    // [18]환적항1, [19]환적항2, [20]모선, [21]프리타임, [22]프리타임 제공 방법, [23]프리타임 제공 만료,
    // [24, 25, 26, 27, 28]도착지비용1{운임포함여부, 항목, 단위, 화폐, 금액} ~ [54, 55, 56, 57, 58]도착지비용7,
    // [59]메모

    for (let i = 2; i < datasheetdata.length; i++) {
      singledata = datasheetdata[i];
    
      // 기준운임1 설정
      let data_absoluteof1 = 0;
      let data_of1 = [
        String(singledata[6]),
        String(singledata[9]),
        String(singledata[12]),
        String(singledata[15])
      ];
      data_of1 = data_of1.filter((item) => item !== "");
      if (data_of1.length > 1 && data_of1.filter((item) => item.includes("+")).length === data_of1.length - 1) {
        data_absoluteof1 = Number(data_of1.filter((item) => !item.includes("+"))[0]);
      } else {
        data_absoluteof1 = 0;
      }

      // 기준운임2 설정
      let data_absoluteof2 = 0;
      let data_of2 = [
        String(singledata[7]),
        String(singledata[10]),
        String(singledata[13]),
        String(singledata[16])
      ];
      data_of2 = data_of2.filter((item) => item !== "");
      if (data_of2.length > 1 && data_of2.filter((item) => item.includes("+")).length === data_of2.length - 1) {
        data_absoluteof2 = Number(data_of2.filter((item) => !item.includes("+"))[0]);
      } else {
        data_absoluteof2 = 0;
      }

      // 기준소요일 설정
      let data_absolutett = "0";
      let data_tt = [
        String(singledata[8]),
        String(singledata[11]),
        String(singledata[14]),
        String(singledata[17])
      ];
      data_tt = data_tt.filter((item) => item !== "");
      if (data_tt.length > 1 && data_tt.filter((item) => item.includes("+")).length === data_tt.length - 1) {
        data_absolutett = data_tt.filter((item) => !item.includes("+"))[0];
      } else {
        data_absolutett = "0";
      }

      // 도착지 정보 확인
      let data_podindex = locationlist.map((row) => row[4] as string).indexOf(singledata[0] as string);
      if (data_podindex === -1) {
        console.log((i + 1) + "행 도착지 오류: " + singledata[0] + "는(은) 유효한 도착지가 아닙니다.");
        continue;
      }

      // 유효기간 확인
      let data_validity =
        (new Date() >= new Date(singledata[3] as string)) && (new Date(singledata[4] as string) >= new Date());
      if (!data_validity) {
        continue;
      }

      // 운임을 제외한 나머지 정보 작성
      singleitem = [
        singledata[1] as string,
        "출발 위치", "POL",
        locationlist[data_podindex][1] as string,
        locationlist[data_podindex][3] as string,
        singledata[0] as string,
        (singledata[18] as string) !== "" ? "환적" : "직항",
        singledata[18] as string,
        singledata[19] as string,
        singledata[2] as string,
        "운송 소요일",
        singledata[3] as string,
        singledata[4] as string,
        "20피트 운임", "40피트 하이큐브 운임", "LCL 운임",
        ((singledata[21] as string) !== "" && (new Date(singledata[23] as string) > new Date())) ? singledata[21] as string : "견적 시 문의",
        ((locationlist[data_podindex][1] as string) === "미국" || (locationlist[data_podindex][1] as string) === "캐나다") ? 1 : 0
      ];

      // 도착지 비용 작성
      let data_addedsurcharge = 0;
      let data_addedsurcharge_20std = 0;
      let data_addedsurcharge_40hc = 0;
      let data_addedsurcharge_lcl = 0;
      let e = 18;
      for (let ii of [24, 29, 34, 39, 44, 49, 54]) {
        // 운임에 포함할 도착지 비용을 변수에 더하고 0으로 만듦
        if (singledata[ii] as boolean) {
          if (singledata[ii + 3] as string !== "USD") {
            let currencyindex = exchangeratelist.map((row) => row[1] as string).indexOf(singledata[ii + 3] as string);
            if (currencyindex === -1) {
              console.log((i + 1) + "행 화폐 단위 오류: " + singledata[ii + 3] + "/USD 환율 정보를 가져올 수 없습니다.");
            continue;
            } else {
            data_addedsurcharge = Math.round((singledata[ii + 4] as number) * exchangeratelist[currencyindex][0] as number);
            data_addedsurcharge = Math.ceil((3.1416 + data_addedsurcharge) / 5) * 5;
            }
          } else {
            data_addedsurcharge = (singledata[ii + 4] as number);
          }
            if (singledata[ii + 2] as string === "CON") {
            data_addedsurcharge_20std += data_addedsurcharge;
            data_addedsurcharge_40hc += data_addedsurcharge;
          } else if (singledata[ii + 2] as string === "TEU") {
            data_addedsurcharge_20std += data_addedsurcharge;
            data_addedsurcharge_40hc += 2 * data_addedsurcharge;
          } else if (singledata[ii + 2] as string === "RT") {
            data_addedsurcharge_lcl += data_addedsurcharge;
          } else if (singledata[ii + 2] as string === "20") {
            data_addedsurcharge_20std += data_addedsurcharge;
          } else if (singledata[ii + 2] as string === "40") {
            data_addedsurcharge_40hc += data_addedsurcharge;
          }
          singledata[ii + 4] = 0;
          data_addedsurcharge = 0;
        }
        singleitem[e] = singledata[ii + 1] as string;
        singleitem[e + 1] = singledata[ii + 2] as string;
        singleitem[e + 2] = singledata[ii + 3] as string;
        singleitem[e + 3] = singledata[ii + 4] as number;
        e += 4;
      }

      // 출발지, 운송 소요일, 운임 작성
      const pollist = [
        ["부산", "BUSAN"],
        ["인천", "INCHEON"],
        ["광양", "GWANGYANG"],
        ["평택-당진", "PYEONGTAEK"],
      ];
      e = 6;
      for (let ii = 0; ii < pollist.length; ii++) {
        // 운임이 없으면 건너뜀
        if (singledata[e] as string === "" && singledata[e + 1] as string === "") {
          e += 3;
          continue;
        }
        // 출발지
        singleitem[1] = pollist[ii][0] as string;
        singleitem[2] = pollist[ii][1] as string;
        // 운송 소요일
        let data_tt_formatted = "";
          // 운송 소요일이 절대값이 아닌 경우
        if (String(singledata[e + 2]).includes("+")) {
            // 기준 소요일이 범위로 지정된 경우
          if (String(data_absolutett).includes("~")) {
            let data_tt_range = data_absolutett.split("~").map((item) => item.trim());
            data_tt_formatted = (
              (Number(data_tt_range[0]) + Number(String(singledata[e + 2]).replace("+", ""))) +
              "~" +
              (Number(data_tt_range[1]) + Number(String(singledata[e + 2]).replace("+", "")))
            );
            // 기준 소요일이 범위로 지정되지 않은 경우
          } else {
            data_tt_formatted = (
              String(Number(data_absolutett) + Number(String(singledata[e + 2]).replace("+", "")))
            );
          }
          // 운송 소요일이 절대값인 경우
        } else if ((singledata[e + 2] as string) !== "") {
          data_tt_formatted = singledata[e + 2] as string;
          // 운송 소요일이 없는 경우
        } else {
          data_tt_formatted = "견적 시 문의";
        }
        singleitem[10] = data_tt_formatted;

        // FCL 운임
        if (singledata[2] === "FCL") {
          // 20피트
          if (String(singledata[e]).includes("+")) {
            // 운임이 절대값이 아닌 경우
            singleitem[13] = Number(String(singledata[e]).replace("+", ""));
            singleitem[13] = Number(singleitem[13]) + data_absoluteof1 + data_addedsurcharge_20std;
            singleitem[13] = Math.max(
              Number(singleitem[13]) * (100 + marginsetting[0][0] as number) / 100,
              Number(singleitem[13]) + marginsetting[0][1] as number
            );
            singleitem[13] = Math.ceil(Number(singleitem[13]) / 10) * 10;
          } else if (Number(singledata[e] as string) !== 0) {
            // 운임이 절대값인 경우
            singleitem[13] = Number(singledata[e]) + data_addedsurcharge_20std;
            singleitem[13] = Math.max(
              Number(singleitem[13]) * (100 + marginsetting[0][0] as number) / 100,
              Number(singleitem[13]) + marginsetting[0][1] as number
            );
            singleitem[13] = Math.ceil(Number(singleitem[13]) / 10) * 10;
            // 운임이 없는 경우
          } else {
            singleitem[13] = "";
          }
          // 40피트 하이큐브
          if (String(singledata[e + 1]).includes("+")) {
            // 운임이 절대값이 아닌 경우
            singleitem[14] = Number(String(singledata[e + 1]).replace("+", ""));
            singleitem[14] = Number(singleitem[14]) + data_absoluteof2 + data_addedsurcharge_40hc;
            singleitem[14] = Math.max(
              Number(singleitem[14]) * (100 + marginsetting[1][0] as number) / 100,
              Number(singleitem[14]) + marginsetting[1][1] as number
            );
            singleitem[14] = Math.ceil(Number(singleitem[14]) / 10) * 10;
          } else if (Number(singledata[e + 1] as string) !== 0) {
            // 운임이 절대값인 경우
            singleitem[14] = Number(singledata[e + 1]) + data_addedsurcharge_40hc;
            singleitem[14] = Math.max(
              Number(singleitem[14]) * (100 + marginsetting[1][0] as number) / 100,
              Number(singleitem[14]) + marginsetting[1][1] as number
            );
            singleitem[14] = Math.ceil(Number(singleitem[14]) / 10) * 10;
            // 운임이 없는 경우
          } else {
            singleitem[14] = "";
          }
          singleitem[15] = "";

          // LCL 운임
        } else {
          singleitem[13] = "";
          singleitem[14] = "";
          if (String(singledata[e]).includes("+")) {
            // 운임이 절대값이 아닌 경우
            singleitem[15] = Number(String(singledata[e]).replace("+", ""));
            singleitem[15] = Number(singleitem[15]) + data_absoluteof1 + data_addedsurcharge_lcl;
            singleitem[15] = Math.max(
              Number(singleitem[15]) * (100 + marginsetting[2][0] as number) / 100,
              Number(singleitem[15]) + marginsetting[2][1] as number
            );
            singleitem[15] = Math.ceil(Number(singleitem[15]) / 10) * 10;
          } else if (Number(singledata[e] as string) !== 0) {
            // 운임이 절대값인 경우
            singleitem[15] = Number(singledata[e]) + data_addedsurcharge_lcl;
            singleitem[15] = Math.max(
              Number(singleitem[15]) * (100 + marginsetting[2][0] as number) / 100,
              Number(singleitem[15]) + marginsetting[2][1] as number
            );
            singleitem[15] = Math.ceil(Number(singleitem[15]) / 10) * 10;
          } // LCL은 운임이 없는 경우를 생략함 (이미 조건에 포함됨)
        }

        deliversheetdata.push([...singleitem]);
        e += 3;
      }
    }

    let deliversheetrange = deliversheet.getRange("A:AT").getUsedRange();
    deliversheetrange.load("values");
    await context.sync();

    // 입력할 데이터가 이미 입력된 데이터보다 많으면 범위를 다시 지정함
    if (deliversheetdata.length > deliversheetrange.values.length) {
      deliversheetrange = deliversheet.getRange("A1:AT" + (deliversheetdata.length));
      deliversheetrange.load("values");
      await context.sync();
    }

    // 입력할 데이터가 이미 입력된 데이터보다 적으면 빈 행을 추가함
    while (deliversheetdata.length < deliversheetrange.values.length) {
      deliversheetdata.push([
        // 46개의 빈 요소 추가
        ...Array(46).fill("")
      ]);
    }

    /** 꼭 오류를 고쳤다 싶으면 범위가 다르대요... 디버깅용
    console.log(deliversheetdata.length);
    console.log(deliversheetdata[0].length);
    console.log(deliversheetrange.values.length);
    console.log(deliversheetrange.values[0].length);
    */
    
    deliversheetrange.values = deliversheetdata;
  });
}

// 데이터 조회
async function explore() {
  await Excel.run(async (context) => {
    const settingsheet = context.workbook.worksheets.getItem("설정");
    const searchsettingsrange = settingsheet.getRange("C6:H9");
    const marginsettingsrange = settingsheet.getRange("C14:D16");
    searchsettingsrange.load("values");
    marginsettingsrange.load("values");

    const datasheet = context.workbook.worksheets.getItem("데이터");
    const datasheetrange = datasheet.getUsedRange();
    datasheetrange.load("values");

    const exchangeratesheet = context.workbook.worksheets.getItem("환율");
    const exchangeratesheetrange = exchangeratesheet.getUsedRange();
    exchangeratesheetrange.load("values");

    const locationsheet = context.workbook.worksheets.getItem("위치");
    const locationsheetrange = locationsheet.getRange("A:F").getUsedRange();
    locationsheetrange.load("values");

    const filtersheet = context.workbook.worksheets.getItem("필터");
    const filtersheetrange = filtersheet.getUsedRange();
    filtersheetrange.load("values");

    const resultsheet = context.workbook.worksheets.getItem("검색");
    const resulttitlerange = resultsheet.getRange("C6:J6");
    const resultdataheaderrange = resultsheet.getRange("M7:T8");
    resulttitlerange.load("values");
    resultdataheaderrange.load("values");

    await context.sync();

    const searchsetting = searchsettingsrange.values;
    let marginsetting = marginsettingsrange.values;
    const datasheetvalues = datasheetrange.values;
    const exchangeratelist = exchangeratesheetrange.values.slice(1, undefined);
    const locationlist = locationsheetrange.values;
    const filterlist = filtersheetrange.values;
    const resultdata = resultdataheaderrange.values;

    let searchsettings = {
      // 검색 조건
      from : searchsetting[0][0] as string,
      fromtype : (
        searchsetting[0][0] as string === "All" ? "all" :
        searchsetting[0][1] as boolean ? "filter" :
        searchsetting[0][2] as boolean ? "country" :
        searchsetting[0][3] as boolean ? "region" :
        "location"
      ),
      to : searchsetting[1][0] as string,
      totype : (
        searchsetting[1][0] as string === "All" ? "all" :
        searchsetting[1][1] as boolean ? "filter" :
        searchsetting[1][2] as boolean ? "country" :
        searchsetting[1][3] as boolean ? "region" :
        "location"
      ),
      carrier : searchsetting[2][0] as string,
      carriertype : (
        searchsetting[2][0] as string === "All" ? "all" :
        searchsetting[2][1] as boolean ? "filter" :
        "specific"
      ),
      volumetype : (
        searchsetting[3][0] as string === "LCL" ? "LCL" :
        searchsetting[3][0] as string === "All" ? "all" :
        "FCL"
      ),
      containtransshippingroute : searchsetting[3][2] as boolean,
      containexpiredfare : searchsetting[3][3] as boolean,
      // 옵션
      addmargin : searchsetting[0][4] as boolean,
      expiredfareonly : searchsetting[1][4] as boolean,
      expiredfreetimeonly : searchsetting[2][4] as boolean,
      conditionzero : searchsetting[3][4] as boolean
    };
    if (searchsettings.conditionzero) {
      searchsettings.fromtype = "all";
      searchsettings.totype = "all";
      searchsettings.carriertype = "all";
      searchsettings.volumetype = "all";
      searchsettings.containtransshippingroute = true;
      searchsettings.containexpiredfare = true;
    }
    // 한글로 입력된 위치는 영어로 변환
    let locationindex = 0;
    if(searchsettings.fromtype === "location") {
      locationindex = locationlist.map((row) => row[3] as string).indexOf(searchsettings.from);
      if (locationindex !== -1) {
        searchsettings.from = locationlist[locationindex][4] as string;
      }
    }
    if(searchsettings.totype === "location") {
      locationindex = locationlist.map((row) => row[3] as string).indexOf(searchsettings.to);
      if (locationindex !== -1) {
        searchsettings.to = locationlist[locationindex][4] as string;
      }
    }

    let searchfor: { from: string[]; to: string[]; carrier: string[] } = {
      from: [],
      to: [],
      carrier: []
    };
    let filterindex = 0;
    if (searchsettings.fromtype === "filter") {
      filterindex = filterlist[1].indexOf(searchsettings.from);
      if (filterindex !== -1) {
        for (let i = 2; i < filterlist.length; i++) {
          if (filterlist[i][filterindex] as string !== "") {
            searchfor.from.push(filterlist[i][filterindex] as string);
          }
        }
      }
    } else if (searchsettings.fromtype === "country") {
      searchfor.from = locationlist.filter((row) => row[1] === searchsettings.from).map((row) => row[4] as string);
    } else if (searchsettings.fromtype === "region") {
      searchfor.from = locationlist.filter((row) => row[2] === searchsettings.from).map((row) => row[4] as string);
    } else if (searchsettings.fromtype === "location") {
      searchfor.from = [searchsettings.from];
    }
    if (searchsettings.totype === "filter") {
      filterindex = filterlist[1].indexOf(searchsettings.to);
      if (filterindex !== -1) {
        for (let i = 2; i < filterlist.length; i++) {
          if (filterlist[i][filterindex] as string !== "") {
            searchfor.to.push(filterlist[i][filterindex] as string);
          }
        }
      }
    } else if (searchsettings.totype === "country") {
      searchfor.to = locationlist.filter((row) => row[1] === searchsettings.to).map((row) => row[4] as string);
    } else if (searchsettings.totype === "region") {
      searchfor.to = locationlist.filter((row) => row[2] === searchsettings.to).map((row) => row[4] as string);
    } else if (searchsettings.totype === "location") {
      searchfor.to = [searchsettings.to];
    }
    if (searchsettings.carriertype === "filter") {
      filterindex = filterlist[1].indexOf(searchsettings.carrier);
      if (filterindex !== -1) {
        for (let i = 2; i < filterlist.length; i++) {
          if (filterlist[i][filterindex] as string !== "") {
            searchfor.carrier.push(filterlist[i][filterindex] as string);
          }
        }
      }
    } else if (searchsettings.carriertype === "specific") {
      searchfor.carrier = [searchsettings.carrier];
    }

    if (!searchsettings.addmargin) {
      marginsetting = [[0, 0], [0, 0], [0, 0]];
    }

    const today = new Date();
    resulttitlerange.values = [[
      locationlist.map((row) => row[4] as string).indexOf(searchsettings.from) === -1 ?
        searchsettings.from : locationlist[locationlist.map((row) => row[4] as string).indexOf(searchsettings.from)][3],
      locationlist.map((row) => row[4] as string).indexOf(searchsettings.to) === -1 ?
        searchsettings.to : locationlist[locationlist.map((row) => row[4] as string).indexOf(searchsettings.to)][3],
      searchsettings.volumetype === "all" ? "모두" : searchsettings.volumetype,
      searchsettings.carriertype === "all" ? "모두" : searchsettings.carrier,
      (today.getFullYear() + "-" + (today.getMonth() + 1) + "-" + today.getDate()),
      searchsettings.addmargin as boolean,
      searchsettings.containtransshippingroute as boolean,
      searchsettings.containexpiredfare as boolean,
    ]];

    let searchresult: (string | number | boolean)[] = [];
    // searchresult 구조: [0]출발 위치, [1]환적항1, [2]환적항2, [3]도착 위치, [4]소요일, [5]운송사, [6]운임, [7]유효기간 종료

    let singledata: (string | number | boolean)[] = [];
    // singledata 구조: [0]도착 위치, [1]운송사, [2]운송 단위, [3]유효기간 시작, [4]유효기간 종료, [5]운임 특성,
    // [6, 7, 8]부산발 정보{운임1, 운임2, 소요일}, [9, 10, 11]인천발 정보, [12, 13, 14]광양발 정보, [15, 16, 17]평택-당진발 정보,
    // [18]환적항1, [19]환적항2, [20]모선, [21]프리타임, [22]프리타임 제공 방법, [23]프리타임 제공 만료,
    // [24, 25, 26, 27, 28]도착지비용1{운임포함여부, 항목, 단위, 화폐, 금액} ~ [54, 55, 56, 57, 58]도착지비용7,
    // [59]메모

    for (let i = 2; i < datasheetvalues.length; i++) {
      singledata = datasheetvalues[i];
      if (
        // 출발지 확인
        (
          (searchfor.from.indexOf("BUSAN") !== -1 && (String(singledata[6]) !== "" || String(singledata[7]) !== "")) ||
          (searchfor.from.indexOf("INCHEON") !== -1 && (String(singledata[9]) !== "" || String(singledata[10]) !== "")) ||
          (searchfor.from.indexOf("GWANGYANG") !== -1 && (String(singledata[12]) !== "" || String(singledata[13]) !== "")) ||
          (searchfor.from.indexOf("PYEONGTAEK") !== -1 && (String(singledata[15]) !== "" || String(singledata[16]) !== "")) ||
          searchsettings.fromtype === "all"
        ) &&
        // 도착지 확인
        (searchfor.to.indexOf(singledata[0] as string) !== -1 || searchsettings.totype === "all" ) &&
        // 운송사 확인
        (searchfor.carrier.indexOf(singledata[1] as string) !== -1 || searchsettings.carriertype === "all") &&
        // 운송 단위 확인
        (singledata[2] as string === searchsettings.volumetype || searchsettings.volumetype === "all") &&
        // 환적 경로 확인
        (singledata[18] as string === "" || searchsettings.containtransshippingroute) &&
        // 만료 데이터 확인
        (new Date(singledata[4] as string) >= new Date() || searchsettings.containexpiredfare || searchsettings.expiredfareonly) &&
        // 운임 만료 데이터만 표시 옵션일 경우 만료 데이터만 포함
        (!searchsettings.expiredfareonly || new Date(singledata[4] as string) < new Date()) &&
        // 프리타임 만료 데이터만 표시 옵션일 경우 만료 데이터만 포함
        (!searchsettings.expiredfreetimeonly || new Date(singledata[23] as string) < new Date())
      ) {
        // 기준운임1 설정
        let data_absoluteof1 = 0;
        let data_of1 = [
          String(singledata[6]),
          String(singledata[9]),
          String(singledata[12]),
          String(singledata[15])
        ];
        data_of1 = data_of1.filter((item) => item !== "");
        if (data_of1.length > 1 && data_of1.filter((item) => item.includes("+")).length === data_of1.length - 1) {
          data_absoluteof1 = Number(data_of1.filter((item) => !item.includes("+"))[0]);
        } else {
          data_absoluteof1 = 0;
        }

        // 기준운임2 설정
        let data_absoluteof2 = 0;
        let data_of2 = [
          String(singledata[7]),
          String(singledata[10]),
          String(singledata[13]),
          String(singledata[16])
        ];
        data_of2 = data_of2.filter((item) => item !== "");
        if (data_of2.length > 1 && data_of2.filter((item) => item.includes("+")).length === data_of2.length - 1) {
          data_absoluteof2 = Number(data_of2.filter((item) => !item.includes("+"))[0]);
        } else {
          data_absoluteof2 = 0;
        }

        // 기준소요일 설정
        let data_absolutett = "0";
        let data_tt = [
          String(singledata[8]),
          String(singledata[11]),
          String(singledata[14]),
          String(singledata[17])
        ];
        data_tt = data_tt.filter((item) => item !== "");
        if (data_tt.length > 1 && data_tt.filter((item) => item.includes("+")).length === data_tt.length - 1) {
          data_absolutett = data_tt.filter((item) => !item.includes("+"))[0];
        } else {
          data_absolutett = "0";
        }

        // 도착지 비용 작성
        let data_addedsurcharge = 0;
        let data_addedsurcharge_20std = 0;
        let data_addedsurcharge_40hc = 0;
        let data_addedsurcharge_lcl = 0;
        for (let ii of [24, 29, 34, 39, 44, 49, 54]) {
          // 운임에 포함할 도착지 비용을 변수에 더함
          if (singledata[ii] as boolean) {
            if (singledata[ii + 3] as string !== "USD") {
              let currencyindex = exchangeratelist.map((row) => row[1] as string).indexOf(singledata[ii + 3] as string);
              if (currencyindex === -1) {
                console.log((i + 1) + "행 화폐 단위 오류: " + singledata[ii + 3] + "/USD 환율 정보를 가져올 수 없습니다.");
              continue;
              } else {
              data_addedsurcharge = Math.round((singledata[ii + 4] as number) * exchangeratelist[currencyindex][0] as number);
              data_addedsurcharge = Math.ceil((3.1416 + data_addedsurcharge) / 5) * 5;
              }
            } else {
              data_addedsurcharge = (singledata[ii + 4] as number);
            }
              if (singledata[ii + 2] as string === "CON") {
              data_addedsurcharge_20std += data_addedsurcharge;
              data_addedsurcharge_40hc += data_addedsurcharge;
            } else if (singledata[ii + 2] as string === "TEU") {
              data_addedsurcharge_20std += data_addedsurcharge;
              data_addedsurcharge_40hc += 2 * data_addedsurcharge;
            } else if (singledata[ii + 2] as string === "RT") {
              data_addedsurcharge_lcl += data_addedsurcharge;
            } else if (singledata[ii + 2] as string === "20") {
              data_addedsurcharge_20std += data_addedsurcharge;
            } else if (singledata[ii + 2] as string === "40") {
              data_addedsurcharge_40hc += data_addedsurcharge;
            }
            data_addedsurcharge = 0;
          }
        }

        // 검색 결과에 추가
        // 환적항
        searchresult[1] = singledata[18];
        searchresult[2] = singledata[19];
        // 도착 위치 (한글로 변환)
        searchresult[3] = singledata[0];
        locationindex = locationlist.map((row) => row[4] as string).indexOf(singledata[0] as string);
        if (locationindex !== -1) {
          searchresult[3] = locationlist[locationindex][3] as string;
        }
        // 운송사
        searchresult[5] = singledata[1];
        // 유효기간 종료
        searchresult[7] = singledata[4];

        // 부산이 출발지 검색 조건에 포함되고 운임이 존재하는 경우 부산발 정보를 검색 결과에 추가 (출발 위치, 소요일, 운임)
        if (
          (searchfor.from.indexOf("BUSAN") !== -1 || searchsettings.fromtype === "all") &&
          (String(singledata[6]) !== "" || String(singledata[7]) !== "")
        ) {
          // 출발 위치
          searchresult[0] = "부산";
          // 운송 소요일
          let data_tt_formatted = "";
          // 운송 소요일이 절대값이 아닌 경우
          if (String(singledata[8]).includes("+")) {
            // 기준 소요일이 범위로 지정된 경우
            if (String(data_absolutett).includes("~")) {
              let data_tt_range = data_absolutett.split("~").map((item) => item.trim());
              data_tt_formatted = (
                (Number(data_tt_range[0]) + Number(String(singledata[8]).replace("+", ""))) +
                "~" +
                (Number(data_tt_range[1]) + Number(String(singledata[8]).replace("+", "")))
              );
            // 기준 소요일이 범위로 지정되지 않은 경우
            } else {
              data_tt_formatted = (
                String(Number(data_absolutett) + Number(String(singledata[8]).replace("+", "")))
              );
            }
            // 운송 소요일이 절대값인 경우
          } else if ((singledata[8] as string) !== "") {
            data_tt_formatted = singledata[8] as string;
            // 운송 소요일이 없는 경우
          } else {
            data_tt_formatted = "견적 시 문의";
          }
          searchresult[4] = data_tt_formatted;
          
          // 운임
          let data_fare_formatted = "";
          let data_fare = 0;
            // FCL 운임
          if (singledata[2] === "FCL") {
            // 20피트
            if (String(singledata[6]).includes("+")) {
              // 운임이 절대값이 아닌 경우
              data_fare = Number(String(singledata[6]).replace("+", ""));
              data_fare = Number(data_fare) + data_absoluteof1 + data_addedsurcharge_20std;
              data_fare = Math.max(
                Number(data_fare) * (100 + marginsetting[0][0] as number) / 100,
                Number(data_fare) + marginsetting[0][1] as number
              );
              data_fare = Math.ceil(Number(data_fare) / 10) * 10;
              data_fare_formatted = String(data_fare);
            } else if (Number(singledata[6] as string) !== 0) {
              // 운임이 절대값인 경우
              data_fare = Number(singledata[6]) + data_addedsurcharge_20std;
              data_fare = Math.max(
                Number(data_fare) * (100 + marginsetting[0][0] as number) / 100,
                Number(data_fare) + marginsetting[0][1] as number
              );
              data_fare = Math.ceil(Number(data_fare) / 10) * 10;
              data_fare_formatted = String(data_fare);
              // 운임이 없는 경우
            } else {
              data_fare = 0;
              data_fare_formatted = "- "
            }
            data_fare_formatted = data_fare_formatted + " | ";
            // 40피트 하이큐브
            if (String(singledata[7]).includes("+")) {
              // 운임이 절대값이 아닌 경우
              data_fare = Number(String(singledata[7]).replace("+", ""));
              data_fare = Number(data_fare) + data_absoluteof2 + data_addedsurcharge_40hc;
              data_fare = Math.max(
                Number(data_fare) * (100 + marginsetting[1][0] as number) / 100,
                Number(data_fare) + marginsetting[1][1] as number
              );
              data_fare = Math.ceil(Number(data_fare) / 10) * 10;
              data_fare_formatted = data_fare_formatted + String(data_fare);
            } else if (Number(singledata[7] as string) !== 0) {
              // 운임이 절대값인 경우
              data_fare = Number(singledata[7]) + data_addedsurcharge_40hc;
              data_fare = Math.max(
                Number(data_fare) * (100 + marginsetting[1][0] as number) / 100,
                Number(data_fare) + marginsetting[1][1] as number
              );
              data_fare = Math.ceil(Number(data_fare) / 10) * 10;
              data_fare_formatted = data_fare_formatted + String(data_fare);
              // 운임이 없는 경우
            } else {
              data_fare = 0;
              data_fare_formatted = data_fare_formatted + " -";
            }
            // LCL 운임
          } else {
            if (String(singledata[6]).includes("+")) {
              // 운임이 절대값이 아닌 경우
              data_fare = Number(String(singledata[6]).replace("+", ""));
              data_fare = Number(data_fare) + data_absoluteof1 + data_addedsurcharge_lcl;
              data_fare = Math.max(
                Number(data_fare) * (100 + marginsetting[2][0] as number) / 100,
                Number(data_fare) + marginsetting[2][1] as number
              );
              data_fare = Math.ceil(Number(data_fare) / 10) * 10;
              data_fare_formatted = String(data_fare);
            } else if (Number(singledata[6] as string) !== 0) {
              // 운임이 절대값인 경우
              data_fare = Number(singledata[6]) + data_addedsurcharge_lcl;
              data_fare = Math.max(
                Number(data_fare) * (100 + marginsetting[2][0] as number) / 100,
                Number(data_fare) + marginsetting[2][1] as number
              );
              data_fare = Math.ceil(Number(data_fare) / 10) * 10;
              data_fare_formatted = String(data_fare);
            }
          }
          searchresult[6] = data_fare_formatted;
          resultdata.push([...searchresult]);
        }

        // 인천이 출발지 검색 조건에 포함되고 운임이 존재하는 경우 인천발 정보를 검색 결과에 추가 (출발 위치, 소요일, 운임)
        if (
          (searchfor.from.indexOf("INCHEON") !== -1 || searchsettings.fromtype === "all") &&
          (String(singledata[9]) !== "" || String(singledata[10]) !== "")
        ) {
          // 출발 위치
          searchresult[0] = "인천";
          // 운송 소요일
          let data_tt_formatted = "";
          // 운송 소요일이 절대값이 아닌 경우
          if (String(singledata[11]).includes("+")) {
            // 기준 소요일이 범위로 지정된 경우
            if (String(data_absolutett).includes("~")) {
              let data_tt_range = data_absolutett.split("~").map((item) => item.trim());
              data_tt_formatted = (
                (Number(data_tt_range[0]) + Number(String(singledata[11]).replace("+", ""))) +
                "~" +
                (Number(data_tt_range[1]) + Number(String(singledata[11]).replace("+", "")))
              );
            // 기준 소요일이 범위로 지정되지 않은 경우
            } else {
              data_tt_formatted = (
                String(Number(data_absolutett) + Number(String(singledata[11]).replace("+", "")))
              );
            }
            // 운송 소요일이 절대값인 경우
          } else if ((singledata[11] as string) !== "") {
            data_tt_formatted = singledata[11] as string;
            // 운송 소요일이 없는 경우
          } else {
            data_tt_formatted = "견적 시 문의";
          }
          searchresult[4] = data_tt_formatted;
          
          // 운임
          let data_fare_formatted = "";
          let data_fare = 0;
            // FCL 운임
          if (singledata[2] === "FCL") {
            // 20피트
            if (String(singledata[9]).includes("+")) {
              // 운임이 절대값이 아닌 경우
              data_fare = Number(String(singledata[9]).replace("+", ""));
              data_fare = Number(data_fare) + data_absoluteof1 + data_addedsurcharge_20std;
              data_fare = Math.max(
                Number(data_fare) * (100 + marginsetting[0][0] as number) / 100,
                Number(data_fare) + marginsetting[0][1] as number
              );
              data_fare = Math.ceil(Number(data_fare) / 10) * 10;
              data_fare_formatted = String(data_fare);
            } else if (Number(singledata[9] as string) !== 0) {
              // 운임이 절대값인 경우
              data_fare = Number(singledata[9]) + data_addedsurcharge_20std;
              data_fare = Math.max(
                Number(data_fare) * (100 + marginsetting[0][0] as number) / 100,
                Number(data_fare) + marginsetting[0][1] as number
              );
              data_fare = Math.ceil(Number(data_fare) / 10) * 10;
              data_fare_formatted = String(data_fare);
              // 운임이 없는 경우
            } else {
              data_fare = 0;
              data_fare_formatted = "- "
            }
            data_fare_formatted = data_fare_formatted + " | ";
            // 40피트 하이큐브
            if (String(singledata[10]).includes("+")) {
              // 운임이 절대값이 아닌 경우
              data_fare = Number(String(singledata[10]).replace("+", ""));
              data_fare = Number(data_fare) + data_absoluteof2 + data_addedsurcharge_40hc;
              data_fare = Math.max(
                Number(data_fare) * (100 + marginsetting[1][0] as number) / 100,
                Number(data_fare) + marginsetting[1][1] as number
              );
              data_fare = Math.ceil(Number(data_fare) / 10) * 10;
              data_fare_formatted = data_fare_formatted + String(data_fare);
            } else if (Number(singledata[10] as string) !== 0) {
              // 운임이 절대값인 경우
              data_fare = Number(singledata[10]) + data_addedsurcharge_40hc;
              data_fare = Math.max(
                Number(data_fare) * (100 + marginsetting[1][0] as number) / 100,
                Number(data_fare) + marginsetting[1][1] as number
              );
              data_fare = Math.ceil(Number(data_fare) / 10) * 10;
              data_fare_formatted = data_fare_formatted + String(data_fare);
              // 운임이 없는 경우
            } else {
              data_fare = 0;
              data_fare_formatted = data_fare_formatted + " -";
            }
            // LCL 운임
          } else {
            if (String(singledata[9]).includes("+")) {
              // 운임이 절대값이 아닌 경우
              data_fare = Number(String(singledata[9]).replace("+", ""));
              data_fare = Number(data_fare) + data_absoluteof1 + data_addedsurcharge_lcl;
              data_fare = Math.max(
                Number(data_fare) * (100 + marginsetting[2][0] as number) / 100,
                Number(data_fare) + marginsetting[2][1] as number
              );
              data_fare = Math.ceil(Number(data_fare) / 10) * 10;
              data_fare_formatted = String(data_fare);
            } else if (Number(singledata[9] as string) !== 0) {
              // 운임이 절대값인 경우
              data_fare = Number(singledata[9]) + data_addedsurcharge_lcl;
              data_fare = Math.max(
                Number(data_fare) * (100 + marginsetting[2][0] as number) / 100,
                Number(data_fare) + marginsetting[2][1] as number
              );
              data_fare = Math.ceil(Number(data_fare) / 10) * 10;
              data_fare_formatted = String(data_fare);
            }
          }
          searchresult[6] = data_fare_formatted;
          resultdata.push([...searchresult]);
        }

        // 광양이 출발지 검색 조건에 포함되고 운임이 존재하는 경우 광양발 정보를 검색 결과에 추가 (출발 위치, 소요일, 운임)
        if (
          (searchfor.from.indexOf("GWANGYANG") !== -1 || searchsettings.fromtype === "all") &&
          (String(singledata[12]) !== "" || String(singledata[13]) !== "")
        ) {
          // 출발 위치
          searchresult[0] = "광양";
          // 운송 소요일
          let data_tt_formatted = "";
          // 운송 소요일이 절대값이 아닌 경우
          if (String(singledata[14]).includes("+")) {
            // 기준 소요일이 범위로 지정된 경우
            if (String(data_absolutett).includes("~")) {
              let data_tt_range = data_absolutett.split("~").map((item) => item.trim());
              data_tt_formatted = (
                (Number(data_tt_range[0]) + Number(String(singledata[14]).replace("+", ""))) +
                "~" +
                (Number(data_tt_range[1]) + Number(String(singledata[14]).replace("+", "")))
              );
            // 기준 소요일이 범위로 지정되지 않은 경우
            } else {
              data_tt_formatted = (
                String(Number(data_absolutett) + Number(String(singledata[14]).replace("+", "")))
              );
            }
            // 운송 소요일이 절대값인 경우
          } else if ((singledata[14] as string) !== "") {
            data_tt_formatted = singledata[14] as string;
            // 운송 소요일이 없는 경우
          } else {
            data_tt_formatted = "견적 시 문의";
          }
          searchresult[4] = data_tt_formatted;
          
          // 운임
          let data_fare_formatted = "";
          let data_fare = 0;
            // FCL 운임
          if (singledata[2] === "FCL") {
            // 20피트
            if (String(singledata[12]).includes("+")) {
              // 운임이 절대값이 아닌 경우
              data_fare = Number(String(singledata[12]).replace("+", ""));
              data_fare = Number(data_fare) + data_absoluteof1 + data_addedsurcharge_20std;
              data_fare = Math.max(
                Number(data_fare) * (100 + marginsetting[0][0] as number) / 100,
                Number(data_fare) + marginsetting[0][1] as number
              );
              data_fare = Math.ceil(Number(data_fare) / 10) * 10;
              data_fare_formatted = String(data_fare);
            } else if (Number(singledata[12] as string) !== 0) {
              // 운임이 절대값인 경우
              data_fare = Number(singledata[12]) + data_addedsurcharge_20std;
              data_fare = Math.max(
                Number(data_fare) * (100 + marginsetting[0][0] as number) / 100,
                Number(data_fare) + marginsetting[0][1] as number
              );
              data_fare = Math.ceil(Number(data_fare) / 10) * 10;
              data_fare_formatted = String(data_fare);
              // 운임이 없는 경우
            } else {
              data_fare = 0;
              data_fare_formatted = "- "
            }
            data_fare_formatted = data_fare_formatted + " | ";
            // 40피트 하이큐브
            if (String(singledata[13]).includes("+")) {
              // 운임이 절대값이 아닌 경우
              data_fare = Number(String(singledata[13]).replace("+", ""));
              data_fare = Number(data_fare) + data_absoluteof2 + data_addedsurcharge_40hc;
              data_fare = Math.max(
                Number(data_fare) * (100 + marginsetting[1][0] as number) / 100,
                Number(data_fare) + marginsetting[1][1] as number
              );
              data_fare = Math.ceil(Number(data_fare) / 10) * 10;
              data_fare_formatted = data_fare_formatted + String(data_fare);
            } else if (Number(singledata[13] as string) !== 0) {
              // 운임이 절대값인 경우
              data_fare = Number(singledata[13]) + data_addedsurcharge_40hc;
              data_fare = Math.max(
                Number(data_fare) * (100 + marginsetting[1][0] as number) / 100,
                Number(data_fare) + marginsetting[1][1] as number
              );
              data_fare = Math.ceil(Number(data_fare) / 10) * 10;
              data_fare_formatted = data_fare_formatted + String(data_fare);
              // 운임이 없는 경우
            } else {
              data_fare = 0;
              data_fare_formatted = data_fare_formatted + " -";
            }
            // LCL 운임
          } else {
            if (String(singledata[12]).includes("+")) {
              // 운임이 절대값이 아닌 경우
              data_fare = Number(String(singledata[12]).replace("+", ""));
              data_fare = Number(data_fare) + data_absoluteof1 + data_addedsurcharge_lcl;
              data_fare = Math.max(
                Number(data_fare) * (100 + marginsetting[2][0] as number) / 100,
                Number(data_fare) + marginsetting[2][1] as number
              );
              data_fare = Math.ceil(Number(data_fare) / 10) * 10;
              data_fare_formatted = String(data_fare);
            } else if (Number(singledata[12] as string) !== 0) {
              // 운임이 절대값인 경우
              data_fare = Number(singledata[12]) + data_addedsurcharge_lcl;
              data_fare = Math.max(
                Number(data_fare) * (100 + marginsetting[2][0] as number) / 100,
                Number(data_fare) + marginsetting[2][1] as number
              );
              data_fare = Math.ceil(Number(data_fare) / 10) * 10;
              data_fare_formatted = String(data_fare);
            }
          }
          searchresult[6] = data_fare_formatted;
          resultdata.push([...searchresult]);
        }

        // 평택-당진이 출발지 검색 조건에 포함되고 운임이 존재하는 경우 평택-당진발 정보를 검색 결과에 추가 (출발 위치, 소요일, 운임)
        if (
          (searchfor.from.indexOf("PYEONGTAEK") !== -1 || searchsettings.fromtype === "all") &&
          (String(singledata[15]) !== "" || String(singledata[16]) !== "")
        ) {
          // 출발 위치
          searchresult[0] = "평택-당진";
          // 운송 소요일
          let data_tt_formatted = "";
          // 운송 소요일이 절대값이 아닌 경우
          if (String(singledata[17]).includes("+")) {
            // 기준 소요일이 범위로 지정된 경우
            if (String(data_absolutett).includes("~")) {
              let data_tt_range = data_absolutett.split("~").map((item) => item.trim());
              data_tt_formatted = (
                (Number(data_tt_range[0]) + Number(String(singledata[17]).replace("+", ""))) +
                "~" +
                (Number(data_tt_range[1]) + Number(String(singledata[17]).replace("+", "")))
              );
            // 기준 소요일이 범위로 지정되지 않은 경우
            } else {
              data_tt_formatted = (
                String(Number(data_absolutett) + Number(String(singledata[17]).replace("+", "")))
              );
            }
            // 운송 소요일이 절대값인 경우
          } else if ((singledata[17] as string) !== "") {
            data_tt_formatted = singledata[17] as string;
            // 운송 소요일이 없는 경우
          } else {
            data_tt_formatted = "견적 시 문의";
          }
          searchresult[4] = data_tt_formatted;
          
          // 운임
          let data_fare_formatted = "";
          let data_fare = 0;
            // FCL 운임
          if (singledata[2] === "FCL") {
            // 20피트
            if (String(singledata[15]).includes("+")) {
              // 운임이 절대값이 아닌 경우
              data_fare = Number(String(singledata[15]).replace("+", ""));
              data_fare = Number(data_fare) + data_absoluteof1 + data_addedsurcharge_20std;
              data_fare = Math.max(
                Number(data_fare) * (100 + marginsetting[0][0] as number) / 100,
                Number(data_fare) + marginsetting[0][1] as number
              );
              data_fare = Math.ceil(Number(data_fare) / 10) * 10;
              data_fare_formatted = String(data_fare);
            } else if (Number(singledata[15] as string) !== 0) {
              // 운임이 절대값인 경우
              data_fare = Number(singledata[15]) + data_addedsurcharge_20std;
              data_fare = Math.max(
                Number(data_fare) * (100 + marginsetting[0][0] as number) / 100,
                Number(data_fare) + marginsetting[0][1] as number
              );
              data_fare = Math.ceil(Number(data_fare) / 10) * 10;
              data_fare_formatted = String(data_fare);
              // 운임이 없는 경우
            } else {
              data_fare = 0;
              data_fare_formatted = "- "
            }
            data_fare_formatted = data_fare_formatted + " | ";
            // 40피트 하이큐브
            if (String(singledata[16]).includes("+")) {
              // 운임이 절대값이 아닌 경우
              data_fare = Number(String(singledata[16]).replace("+", ""));
              data_fare = Number(data_fare) + data_absoluteof2 + data_addedsurcharge_40hc;
              data_fare = Math.max(
                Number(data_fare) * (100 + marginsetting[1][0] as number) / 100,
                Number(data_fare) + marginsetting[1][1] as number
              );
              data_fare = Math.ceil(Number(data_fare) / 10) * 10;
              data_fare_formatted = data_fare_formatted + String(data_fare);
            } else if (Number(singledata[16] as string) !== 0) {
              // 운임이 절대값인 경우
              data_fare = Number(singledata[16]) + data_addedsurcharge_40hc;
              data_fare = Math.max(
                Number(data_fare) * (100 + marginsetting[1][0] as number) / 100,
                Number(data_fare) + marginsetting[1][1] as number
              );
              data_fare = Math.ceil(Number(data_fare) / 10) * 10;
              data_fare_formatted = data_fare_formatted + String(data_fare);
              // 운임이 없는 경우
            } else {
              data_fare = 0;
              data_fare_formatted = data_fare_formatted + " -";
            }
            // LCL 운임
          } else {
            if (String(singledata[15]).includes("+")) {
              // 운임이 절대값이 아닌 경우
              data_fare = Number(String(singledata[15]).replace("+", ""));
              data_fare = Number(data_fare) + data_absoluteof1 + data_addedsurcharge_lcl;
              data_fare = Math.max(
                Number(data_fare) * (100 + marginsetting[2][0] as number) / 100,
                Number(data_fare) + marginsetting[2][1] as number
              );
              data_fare = Math.ceil(Number(data_fare) / 10) * 10;
              data_fare_formatted = String(data_fare);
            } else if (Number(singledata[15] as string) !== 0) {
              // 운임이 절대값인 경우
              data_fare = Number(singledata[15]) + data_addedsurcharge_lcl;
              data_fare = Math.max(
                Number(data_fare) * (100 + marginsetting[2][0] as number) / 100,
                Number(data_fare) + marginsetting[2][1] as number
              );
              data_fare = Math.ceil(Number(data_fare) / 10) * 10;
              data_fare_formatted = String(data_fare);
            }
          }
          searchresult[6] = data_fare_formatted;
          resultdata.push([...searchresult]);
        }
      }
    }

    // 찾은 데이터 수
    resultdata[0][1] = resultdata.length - 2;

    let resultsheetrange = resultsheet.getRange("M:T").getUsedRange();
    resultsheetrange.load("values");
    await context.sync();

    // 입력할 데이터가 이미 입력된 데이터보다 많으면 범위를 다시 지정함
    if (resultdata.length > resultsheetrange.values.length) {
      resultsheetrange = resultsheet.getRange("M7:T" + (resultdata.length + 6));
      resultsheetrange.load("values");
      await context.sync();
    }

    // 입력할 데이터가 이미 입력된 데이터보다 적으면 빈 행을 추가함
    while (resultdata.length < resultsheetrange.values.length) {
      resultdata.push([
        // 8개의 빈 요소 추가
        ...Array(8).fill("")
      ]);
    }

    /** 꼭 오류를 고쳤다 싶으면 범위가 다르대요... 디버깅용
    console.log(deliversheetdata.length);
    console.log(deliversheetdata[0].length);
    console.log(deliversheetrange.values.length);
    console.log(deliversheetrange.values[0].length);
    */

    resultsheetrange.values = resultdata;
  });
}

// 데이터 복구
async function fix() {
  await Excel.run(async (context) => {
    const datasheetrange = context.workbook.worksheets.getItem("데이터").getUsedRange();
    datasheetrange.load("values");

    const logsheetrange = context.workbook.worksheets.getItem("로그").getUsedRange();
    logsheetrange.load("values");

    await context.sync();

    let datasheetdata = datasheetrange.values;
    let logsheetdata = logsheetrange.values;
    let selectedloglist = logsheetdata.slice(1, undefined).filter((row) => row[17] === true);
    if (selectedloglist.length === 0) { return; }

    let singlelog: (string | number | boolean)[] = [];
    // singlelog 구조: [0]날짜와 시간 [1]사용자 [2]유형 [3]도착지 [4]운송사 [5]운송 단위 [6]변경한 항목 [7~11]변경 전 값 [12~16]변경 후 값 [17](선택)
    //
    // 로그 유형의 카테고리:
    // 수정, 신규 (총 2개)
    //
    // 변경한 항목의 카테고리:
    // 전체, 유효 기간, 운임 특성, 부산발 정보, 인천발 정보, 광양발 정보, 평택-당진발 정보, 경로 정보, 메모, 프리타임, 도착지 비용1~7 (총 17개)

    for (let i = 0; i < selectedloglist.length; i++) {
      singlelog = selectedloglist[i];
      let dataindex = datasheetdata.map((row) => (
        row[0] === singlelog[3] &&
        row[1] === singlelog[4] &&
        row[2] === singlelog[5]
      )).indexOf(true);
      if (dataindex === -1) {
        console.log((i + 1) + "행 오류: 데이터 시트에서 해당 데이터를 찾을 수 없습니다.");
        return;
      }
      if (singlelog[2] === "신규") {
        datasheetdata[dataindex] = ([singlelog[3], singlelog[4], singlelog[5]] as string[]).concat(Array(57).fill(""));
      } else if (singlelog[2] === "수정") {
        if (singlelog[6] === "유효 기간") {
          datasheetdata[dataindex][3] = singlelog[7];
          datasheetdata[dataindex][4] = singlelog[8];
        } else if (singlelog[6] === "운임 특성") {
          datasheetdata[dataindex][5] = singlelog[7];
        } else if (singlelog[6] === "부산발 정보") {
          datasheetdata[dataindex][6] = singlelog[7];
          datasheetdata[dataindex][7] = singlelog[8];
          datasheetdata[dataindex][8] = singlelog[9];
        } else if (singlelog[6] === "인천발 정보") {
          datasheetdata[dataindex][9] = singlelog[7];
          datasheetdata[dataindex][10] = singlelog[8];
          datasheetdata[dataindex][11] = singlelog[9];
        } else if (singlelog[6] === "광양발 정보") {
          datasheetdata[dataindex][12] = singlelog[7];
          datasheetdata[dataindex][13] = singlelog[8];
          datasheetdata[dataindex][14] = singlelog[9];
        } else if (singlelog[6] === "평택-당진발 정보") {
          datasheetdata[dataindex][15] = singlelog[7];
          datasheetdata[dataindex][16] = singlelog[8];
          datasheetdata[dataindex][17] = singlelog[9];
        } else if (singlelog[6] === "경로 정보") {
          datasheetdata[dataindex][18] = singlelog[7];
          datasheetdata[dataindex][19] = singlelog[8];
          datasheetdata[dataindex][20] = singlelog[9];
        } else if (singlelog[6] === "프리타임") {
          datasheetdata[dataindex][21] = singlelog[7];
          datasheetdata[dataindex][22] = singlelog[8];
          datasheetdata[dataindex][23] = singlelog[9];
        } else if (singlelog[6] === "도착지 비용1") {
          datasheetdata[dataindex][24] = singlelog[7] === "true";
          datasheetdata[dataindex][25] = singlelog[8];
          datasheetdata[dataindex][26] = singlelog[9];
          datasheetdata[dataindex][27] = singlelog[10];
          datasheetdata[dataindex][28] = Number(singlelog[11]);
        } else if (singlelog[6] === "도착지 비용2") {
          datasheetdata[dataindex][29] = singlelog[7] === "true";
          datasheetdata[dataindex][30] = singlelog[8];
          datasheetdata[dataindex][31] = singlelog[9];
          datasheetdata[dataindex][32] = singlelog[10];
          datasheetdata[dataindex][33] = Number(singlelog[11]);
        } else if (singlelog[6] === "도착지 비용3") {
          datasheetdata[dataindex][34] = singlelog[7] === "true";
          datasheetdata[dataindex][35] = singlelog[8];
          datasheetdata[dataindex][36] = singlelog[9];
          datasheetdata[dataindex][37] = singlelog[10];
          datasheetdata[dataindex][38] = Number(singlelog[11]);
        } else if (singlelog[6] === "도착지 비용4") {
          datasheetdata[dataindex][39] = singlelog[7] === "true";
          datasheetdata[dataindex][40] = singlelog[8];
          datasheetdata[dataindex][41] = singlelog[9];
          datasheetdata[dataindex][42] = singlelog[10];
          datasheetdata[dataindex][43] = Number(singlelog[11]);
        } else if (singlelog[6] === "도착지 비용5") {
          datasheetdata[dataindex][44] = singlelog[7] === "true";
          datasheetdata[dataindex][45] = singlelog[8];
          datasheetdata[dataindex][46] = singlelog[9];
          datasheetdata[dataindex][47] = singlelog[10];
          datasheetdata[dataindex][48] = Number(singlelog[11]);
        } else if (singlelog[6] === "도착지 비용6") {
          datasheetdata[dataindex][49] = singlelog[7] === "true";
          datasheetdata[dataindex][50] = singlelog[8];
          datasheetdata[dataindex][51] = singlelog[9];
          datasheetdata[dataindex][52] = singlelog[10];
          datasheetdata[dataindex][53] = Number(singlelog[11]);
        } else if (singlelog[6] === "도착지 비용7") {
          datasheetdata[dataindex][54] = singlelog[7] === "true";
          datasheetdata[dataindex][55] = singlelog[8];
          datasheetdata[dataindex][56] = singlelog[9];
          datasheetdata[dataindex][57] = singlelog[10];
          datasheetdata[dataindex][58] = Number(singlelog[11]);
        } else if (singlelog[6] === "메모") {
          datasheetdata[dataindex][59] = singlelog[7];
        }
      }
    }

    datasheetrange.values = datasheetdata;

    logsheetdata = logsheetdata.filter((row) => selectedloglist.indexOf(row) === -1);
    logsheetrange.values = logsheetdata.concat(Array(selectedloglist.length).fill(Array(18).fill("")));
    logsheetrange.getColumn(17).control = {type: Excel.CellControlType.empty};

    const logsheetnewrange = context.workbook.worksheets.getItem("로그").getRange("A1:R" + logsheetdata.length);
    logsheetnewrange.getColumn(17).control = {type: Excel.CellControlType.checkbox};
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}
