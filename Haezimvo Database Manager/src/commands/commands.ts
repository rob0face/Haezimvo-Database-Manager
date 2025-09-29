// Type definitions for Office.js are provided by the @types/office-js package.

// Register each function with Office actions
Office.actions.associate("apply", apply);
Office.actions.associate("bring", bring);
Office.actions.associate("combine", combine);
Office.actions.associate("deliver", deliver);
Office.actions.associate("explore", explore);
Office.actions.associate("fix", fix);

// 데이터 반영
async function apply() {
  await Excel.run(async (context) => {
    const applysheet = context.workbook.worksheets.getItem("반영");
    const applysheetrange = applysheet.getRange("C6:K20");
    const portsofladingrange = applysheet.getRange("B12:B15");
    applysheetrange.load("values");
    portsofladingrange.load("values");

    const settingsheet = context.workbook.worksheets.getItem("설정");
    const usersettingrange = settingsheet.getRange("H13:K16");
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
    const portsoflading = portsofladingrange.values.map((row) => row[0] as string);
    const logsheetdata = logsheetrange.values;

    // 새 데이터 작성:
    let newdata: (string | number | boolean)[] = [];

    // 새 로그 작성:
    let newlog: (string | number | boolean)[] = [];
    // newlog 구조:
    // [0]날짜와 시간, [1]사용자, [2]변경 유형, [3]도착지, [4]운송사, [5]운송 단위, [6]변경한 항목, [7~11]변경 전 값, [12~16]변경 후 값, [17](선택)
    //
    // 변경한 항목의 카테고리:
    // 유효 기간, 운임 특성, 선적항1, 선적항2, 선적항3, 선적항4, 경로 정보, 메모, 프리타임, 도착지 비용1~7 (총 16개)

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
        if ([0, 2, 4, 6, 7, 9, 11].indexOf(targetmonth as number) !== -1) {
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

    // 운임 및 소요일 정보
    newdata = newdata.concat(portsoflading[0], applysheetdata[6][0] as string, applysheetdata[6][1] as string, applysheetdata[6][2] as string);
    newdata = newdata.concat(portsoflading[1], applysheetdata[7][0] as string, applysheetdata[7][1] as string, applysheetdata[7][2] as string);
    newdata = newdata.concat(portsoflading[2], applysheetdata[8][0] as string, applysheetdata[8][1] as string, applysheetdata[8][2] as string);
    newdata = newdata.concat(portsoflading[3], applysheetdata[9][0] as string, applysheetdata[9][1] as string, applysheetdata[9][2] as string);

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
    const datasheetinputrange = datasheet.getRange("A" + datasheetinputindex + ":BL" + datasheetinputindex);
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
        datasheetinputrange.values[0][8] !== newdata[8] ||
        datasheetinputrange.values[0][9] !== newdata[9]
      ) {
        newlog[6] = "선적항1 정보";
        newlog[7] = String(datasheetinputrange.values[0][6]);
        newlog[8] = String(datasheetinputrange.values[0][7]);
        newlog[9] = String(datasheetinputrange.values[0][8]);
        newlog[10] = String(datasheetinputrange.values[0][9]);
        newlog[11] = "";
        newlog[12] = String(newdata[6]);
        newlog[13] = String(newdata[7]);
        newlog[14] = String(newdata[8]);
        newlog[15] = String(newdata[9]);
        newlog[16] = "";
        logsheetdata.push([...newlog]);
      }
      if (
        datasheetinputrange.values[0][10] !== newdata[10] || 
        datasheetinputrange.values[0][11] !== newdata[11] ||
        datasheetinputrange.values[0][12] !== newdata[12] ||
        datasheetinputrange.values[0][13] !== newdata[13]
      ) {
        newlog[6] = "선적항2 정보";
        newlog[7] = String(datasheetinputrange.values[0][10]);
        newlog[8] = String(datasheetinputrange.values[0][11]);
        newlog[9] = String(datasheetinputrange.values[0][12]);
        newlog[10] = String(datasheetinputrange.values[0][13]);
        newlog[11] = "";
        newlog[12] = String(newdata[10]);
        newlog[13] = String(newdata[11]);
        newlog[14] = String(newdata[12]);
        newlog[15] = String(newdata[13]);
        newlog[16] = "";
        logsheetdata.push([...newlog]);
      }
      if (
        datasheetinputrange.values[0][14] !== newdata[14] || 
        datasheetinputrange.values[0][15] !== newdata[15] || 
        datasheetinputrange.values[0][16] !== newdata[16] ||
        datasheetinputrange.values[0][17] !== newdata[17]
      ) {
        newlog[6] = "선적항3 정보";
        newlog[7] = String(datasheetinputrange.values[0][14]);
        newlog[8] = String(datasheetinputrange.values[0][15]);
        newlog[9] = String(datasheetinputrange.values[0][16]);
        newlog[10] = String(datasheetinputrange.values[0][17]);
        newlog[11] = "";
        newlog[12] = String(newdata[14]);
        newlog[13] = String(newdata[15]);
        newlog[14] = String(newdata[16]);
        newlog[15] = String(newdata[17]);
        newlog[16] = "";
        logsheetdata.push([...newlog]);
      }
      if (
        datasheetinputrange.values[0][18] !== newdata[18] ||
        datasheetinputrange.values[0][19] !== newdata[19] ||
        datasheetinputrange.values[0][20] !== newdata[20] ||
        datasheetinputrange.values[0][21] !== newdata[21] 
      ) {
        newlog[6] = "선적항4 정보";
        newlog[7] = String(datasheetinputrange.values[0][18]);
        newlog[8] = String(datasheetinputrange.values[0][19]);
        newlog[9] = String(datasheetinputrange.values[0][20]);
        newlog[10] = String(datasheetinputrange.values[0][21]);
        newlog[11] = "";
        newlog[12] = String(newdata[18]);
        newlog[13] = String(newdata[19]);
        newlog[14] = String(newdata[20]);
        newlog[15] = String(newdata[21]);
        newlog[16] = "";
        logsheetdata.push([...newlog]);
      }
      if (
        datasheetinputrange.values[0][22] !== newdata[22] ||
        datasheetinputrange.values[0][23] !== newdata[23] ||
        datasheetinputrange.values[0][24] !== newdata[24]
      ) {
        newlog[6] = "경로 정보";
        newlog[7] = String(datasheetinputrange.values[0][22]);
        newlog[8] = String(datasheetinputrange.values[0][23]);
        newlog[9] = String(datasheetinputrange.values[0][24]);
        newlog[10] = "";
        newlog[11] = "";
        newlog[12] = String(newdata[22]);
        newlog[13] = String(newdata[23]);
        newlog[14] = String(newdata[24]);
        newlog[15] = "";
        newlog[16] = "";
        logsheetdata.push([...newlog]);
      }
      if (
        datasheetinputrange.values[0][25] !== newdata[25] ||
        datasheetinputrange.values[0][26] !== newdata[26] ||
        datasheetinputrange.values[0][27] !== newdata[27]
      ) {
        newlog[6] = "프리타임";
        newlog[7] = String(datasheetinputrange.values[0][25]);
        newlog[8] = String(datasheetinputrange.values[0][26]);
        newlog[9] = String(datasheetinputrange.values[0][27]);
        newlog[10] = "";
        newlog[11] = "";
        newlog[12] = String(newdata[25]);
        newlog[13] = String(newdata[26]);
        newlog[14] = String(newdata[27]);
        newlog[15] = "";
        newlog[16] = "";
        logsheetdata.push([...newlog]);
      }
      for (let i = 1; i <= 7; i++) {
        if (
          datasheetinputrange.values[0][28 + ((i - 1) * 5)] !== newdata[28 + ((i - 1) * 5)] ||
          datasheetinputrange.values[0][29 + ((i - 1) * 5)] !== newdata[29 + ((i - 1) * 5)] ||
          datasheetinputrange.values[0][30 + ((i - 1) * 5)] !== newdata[30 + ((i - 1) * 5)] ||
          datasheetinputrange.values[0][31 + ((i - 1) * 5)] !== newdata[31 + ((i - 1) * 5)] ||
          datasheetinputrange.values[0][32 + ((i - 1) * 5)] !== newdata[32 + ((i - 1) * 5)]
        ) {
          newlog[6] = "도착지 비용" + i;
          newlog[7] = String(datasheetinputrange.values[0][28 + ((i - 1) * 5)]);
          newlog[8] = String(datasheetinputrange.values[0][29 + ((i - 1) * 5)]);
          newlog[9] = String(datasheetinputrange.values[0][30 + ((i - 1) * 5)]);
          newlog[10] = String(datasheetinputrange.values[0][31 + ((i - 1) * 5)]);
          newlog[11] = String(datasheetinputrange.values[0][32 + ((i - 1) * 5)]);
          newlog[12] = String(newdata[28 + ((i - 1) * 5)]);
          newlog[13] = String(newdata[29 + ((i - 1) * 5)]);
          newlog[14] = String(newdata[30 + ((i - 1) * 5)]);
          newlog[15] = String(newdata[31 + ((i - 1) * 5)]);
          newlog[16] = String(newdata[32 + ((i - 1) * 5)]);
          logsheetdata.push([...newlog]);
        }
      }
      if (datasheetinputrange.values[0][63] !== newdata[63]) {
        newlog[6] = "메모";
        newlog[7] = String(datasheetinputrange.values[0][63]);
        newlog[8] = "";
        newlog[9] = "";
        newlog[10] = "";
        newlog[11] = "";
        newlog[12] = String(newdata[63]);
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
    const portsofladingrange = applysheet.getRange("B12:B15");
    const applysheetformularange1 = applysheet.getRange("G6:J6");
    const applysheetformularange2 = applysheet.getRange("C11:D11");
    applysheetrange.load("values");
    portsofladingrange.load("values");
    applysheetformularange1.load("formulas");
    applysheetformularange2.load("formulas");

    const datasheet = context.workbook.worksheets.getItem("데이터");
    const datasheetindexrange = datasheet.getRange("A:C").getUsedRange();
    datasheetindexrange.load("values");

    await context.sync();

    const applysheetdata = applysheetrange.values;
    const portsoflading = portsofladingrange.values;
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

    const datasheetrange = datasheet.getRange("A" + datasheetindex + ":BL" + datasheetindex);
    datasheetrange.load("values");
    await context.sync();
    const datasheetdata = (datasheetrange.values)[0];

    if (datasheetdata[0] as string === "") {
      applysheetdata[3][2] = "신규";
    } else {
      // 유효 기간 부터, 까지, 운임 특성
      applysheetdata[3][0] = datasheetdata[3] as string;
      applysheetdata[3][1] = datasheetdata[4] as string;
      applysheetdata[3][2] = datasheetdata[5] as string;
      // 부산발, 인천발, 광양발, 평택-당진발 운임 및 소요일 정보
      portsoflading[0][0] = datasheetdata[6] as string;
      applysheetdata[6][0] = datasheetdata[7] as string;
      applysheetdata[6][1] = datasheetdata[8] as string;
      applysheetdata[6][2] = datasheetdata[9] as string;
      portsoflading[1][0] = datasheetdata[10] as string;
      applysheetdata[7][0] = datasheetdata[11] as string;
      applysheetdata[7][1] = datasheetdata[12] as string;
      applysheetdata[7][2] = datasheetdata[13] as string;
      portsoflading[2][0] = datasheetdata[14] as string;
      applysheetdata[8][0] = datasheetdata[15] as string;
      applysheetdata[8][1] = datasheetdata[16] as string;
      applysheetdata[8][2] = datasheetdata[17] as string;
      portsoflading[3][0] = datasheetdata[18] as string;
      applysheetdata[9][0] = datasheetdata[19] as string;
      applysheetdata[9][1] = datasheetdata[20] as string;
      applysheetdata[9][2] = datasheetdata[21] as string;
      // 경로 정보
      applysheetdata[12][0] = datasheetdata[22] as string;
      applysheetdata[12][1] = datasheetdata[23] as string;
      applysheetdata[12][2] = datasheetdata[24] as string;
      // 도착 위치 프리타임
      applysheetdata[3][4] = datasheetdata[25] as string;
      applysheetdata[3][5] = datasheetdata[26] as string;
      applysheetdata[3][6] = datasheetdata[27] as string;
      // 도착지 비용 1~7
      applysheetdata[6][4] = datasheetdata[28] as boolean;
      applysheetdata[6][5] = datasheetdata[29] as string;
      applysheetdata[6][6] = datasheetdata[30] as string;
      applysheetdata[6][7] = datasheetdata[31] as string;
      applysheetdata[6][8] = datasheetdata[32] as number;
      applysheetdata[7][4] = datasheetdata[33] as boolean;
      applysheetdata[7][5] = datasheetdata[34] as string;
      applysheetdata[7][6] = datasheetdata[35] as string;
      applysheetdata[7][7] = datasheetdata[36] as string;
      applysheetdata[7][8] = datasheetdata[37] as number;
      applysheetdata[8][4] = datasheetdata[38] as boolean;
      applysheetdata[8][5] = datasheetdata[39] as string;
      applysheetdata[8][6] = datasheetdata[40] as string;
      applysheetdata[8][7] = datasheetdata[41] as string;
      applysheetdata[8][8] = datasheetdata[42] as number;
      applysheetdata[9][4] = datasheetdata[43] as boolean;
      applysheetdata[9][5] = datasheetdata[44] as string;
      applysheetdata[9][6] = datasheetdata[45] as string;
      applysheetdata[9][7] = datasheetdata[46] as string;
      applysheetdata[9][8] = datasheetdata[47] as number;
      applysheetdata[10][4] = datasheetdata[48] as boolean;
      applysheetdata[10][5] = datasheetdata[49] as string;
      applysheetdata[10][6] = datasheetdata[50] as string;
      applysheetdata[10][7] = datasheetdata[51] as string;
      applysheetdata[10][8] = datasheetdata[52] as number;
      applysheetdata[11][4] = datasheetdata[53] as boolean;
      applysheetdata[11][5] = datasheetdata[54] as string;
      applysheetdata[11][6] = datasheetdata[55] as string;
      applysheetdata[11][7] = datasheetdata[56] as string;
      applysheetdata[11][8] = datasheetdata[57] as number;
      applysheetdata[12][4] = datasheetdata[58] as boolean;
      applysheetdata[12][5] = datasheetdata[59] as string;
      applysheetdata[12][6] = datasheetdata[60] as string;
      applysheetdata[12][7] = datasheetdata[61] as string;
      applysheetdata[12][8] = datasheetdata[62] as number;
      // 메모
      applysheetdata[14][0] = datasheetdata[63] as string;
    }

    applysheetrange.values = applysheetdata;
    portsofladingrange.values = portsoflading;

    // 수식 덮어쓰기
    applysheetformularange1.formulas = applysheetformula1;
    applysheetformularange2.formulas = applysheetformula2;
  });
}

// 데이터 병합
async function combine() {
  await Excel.run(async (context) => {
    const settingsheet = context.workbook.worksheets.getItem("설정");
    const usersettingrange = settingsheet.getRange("H13:K16");
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
          datasheetdata[pushindex][8] !== singlepusheddata[8] ||
          datasheetdata[pushindex][9] !== singlepusheddata[9]
        ) {
          newlog[6] = "선적항1 정보";
          newlog[7] = String(datasheetdata[pushindex][6]);
          newlog[8] = String(datasheetdata[pushindex][7]);
          newlog[9] = String(datasheetdata[pushindex][8]);
          newlog[10] = String(datasheetdata[pushindex][9]);
          newlog[11] = "";
          newlog[12] = String(singlepusheddata[6]);
          newlog[13] = String(singlepusheddata[7]);
          newlog[14] = String(singlepusheddata[8]);
          newlog[15] = String(singlepusheddata[9]);
          newlog[16] = "";
          logsheetdata.push([...newlog]);
        }
        if (
          datasheetdata[pushindex][10] !== singlepusheddata[10] ||
          datasheetdata[pushindex][11] !== singlepusheddata[11] ||
          datasheetdata[pushindex][12] !== singlepusheddata[12] ||
          datasheetdata[pushindex][13] !== singlepusheddata[13]
        ) {
          newlog[6] = "선적항2 정보";
          newlog[7] = String(datasheetdata[pushindex][10]);
          newlog[8] = String(datasheetdata[pushindex][11]);
          newlog[9] = String(datasheetdata[pushindex][12]);
          newlog[10] = String(datasheetdata[pushindex][13]);
          newlog[11] = "";
          newlog[12] = String(singlepusheddata[10]);
          newlog[13] = String(singlepusheddata[11]);
          newlog[14] = String(singlepusheddata[12]);
          newlog[15] = String(singlepusheddata[13]);
          newlog[16] = "";
          logsheetdata.push([...newlog]);
        }
        if (
          datasheetdata[pushindex][14] !== singlepusheddata[14] ||
          datasheetdata[pushindex][15] !== singlepusheddata[15] ||
          datasheetdata[pushindex][16] !== singlepusheddata[16] ||
          datasheetdata[pushindex][17] !== singlepusheddata[17]
        ) {
          newlog[6] = "선적항3 정보";
          newlog[7] = String(datasheetdata[pushindex][14]);
          newlog[8] = String(datasheetdata[pushindex][15]);
          newlog[9] = String(datasheetdata[pushindex][16]);
          newlog[10] = String(datasheetdata[pushindex][17]);
          newlog[11] = "";
          newlog[12] = String(singlepusheddata[14]);
          newlog[13] = String(singlepusheddata[15]);
          newlog[14] = String(singlepusheddata[16]);
          newlog[15] = String(singlepusheddata[17]);
          newlog[16] = "";
          logsheetdata.push([...newlog]);
        }
        if (
          datasheetdata[pushindex][18] !== singlepusheddata[18] ||
          datasheetdata[pushindex][19] !== singlepusheddata[19] ||
          datasheetdata[pushindex][20] !== singlepusheddata[20] ||
          datasheetdata[pushindex][21] !== singlepusheddata[21] 
        ) {
          newlog[6] = "선적항4 정보";
          newlog[7] = String(datasheetdata[pushindex][18]);
          newlog[8] = String(datasheetdata[pushindex][19]);
          newlog[9] = String(datasheetdata[pushindex][20]);
          newlog[10] = String(datasheetdata[pushindex][21]);
          newlog[11] = "";
          newlog[12] = String(singlepusheddata[18]);
          newlog[13] = String(singlepusheddata[19]);
          newlog[14] = String(singlepusheddata[20]);
          newlog[15] = String(singlepusheddata[21]);
          newlog[16] = "";
          logsheetdata.push([...newlog]);
        }
        if (
          datasheetdata[pushindex][22] !== singlepusheddata[22] ||
          datasheetdata[pushindex][23] !== singlepusheddata[23] ||
          datasheetdata[pushindex][24] !== singlepusheddata[24]
        ) {
          newlog[6] = "경로 정보";
          newlog[7] = String(datasheetdata[pushindex][22]);
          newlog[8] = String(datasheetdata[pushindex][23]);
          newlog[9] = String(datasheetdata[pushindex][24]);
          newlog[10] = "";
          newlog[11] = "";
          newlog[12] = String(singlepusheddata[22]);
          newlog[13] = String(singlepusheddata[23]);
          newlog[14] = String(singlepusheddata[24]);
          newlog[15] = "";
          newlog[16] = "";
          logsheetdata.push([...newlog]);
        }
        if (
          datasheetdata[pushindex][25] !== singlepusheddata[25] ||
          datasheetdata[pushindex][26] !== singlepusheddata[26] ||
          datasheetdata[pushindex][27] !== singlepusheddata[27]
        ) {
          newlog[6] = "프리타임";
          newlog[7] = String(datasheetdata[pushindex][25]);
          newlog[8] = String(datasheetdata[pushindex][26]);
          newlog[9] = String(datasheetdata[pushindex][27]);
          newlog[10] = "";
          newlog[11] = "";
          newlog[12] = String(singlepusheddata[25]);
          newlog[13] = String(singlepusheddata[26]);
          newlog[14] = String(singlepusheddata[27]);
          newlog[15] = "";
          newlog[16] = "";
          logsheetdata.push([...newlog]);
        }
        for (let ii = 1; ii <= 7; ii++) {
          if (
            datasheetdata[pushindex][28 + ((ii - 1) * 5)] !== singlepusheddata[28 + ((ii - 1) * 5)] ||
            datasheetdata[pushindex][29 + ((ii - 1) * 5)] !== singlepusheddata[29 + ((ii - 1) * 5)] ||
            datasheetdata[pushindex][30 + ((ii - 1) * 5)] !== singlepusheddata[30 + ((ii - 1) * 5)] ||
            datasheetdata[pushindex][31 + ((ii - 1) * 5)] !== singlepusheddata[31 + ((ii - 1) * 5)] ||
            datasheetdata[pushindex][32 + ((ii - 1) * 5)] !== singlepusheddata[32 + ((ii - 1) * 5)]
          ) {
            newlog[6] = "도착지 비용" + ii;
            newlog[7] = String(datasheetdata[pushindex][28 + ((ii - 1) * 5)]);
            newlog[8] = String(datasheetdata[pushindex][29 + ((ii - 1) * 5)]);
            newlog[9] = String(datasheetdata[pushindex][30 + ((ii - 1) * 5)]);
            newlog[10] = String(datasheetdata[pushindex][31 + ((ii - 1) * 5)]);
            newlog[11] = String(datasheetdata[pushindex][32 + ((ii - 1) * 5)]);
            newlog[12] = String(singlepusheddata[28 + ((ii - 1) * 5)]);
            newlog[13] = String(singlepusheddata[29 + ((ii - 1) * 5)]);
            newlog[14] = String(singlepusheddata[30 + ((ii - 1) * 5)]);
            newlog[15] = String(singlepusheddata[31 + ((ii - 1) * 5)]);
            newlog[16] = String(singlepusheddata[32 + ((ii - 1) * 5)]);
            logsheetdata.push([...newlog]);
          }
        }
        if (datasheetdata[pushindex][63] !== singlepusheddata[63]) {
          newlog[6] = "메모";
          newlog[7] = String(datasheetdata[pushindex][63]);
          newlog[8] = "";
          newlog[9] = "";
          newlog[10] = "";
          newlog[11] = "";
          newlog[12] = String(singlepusheddata[63]);
          newlog[13] = "";
          newlog[14] = "";
          newlog[15] = "";
          newlog[16] = "";
          logsheetdata.push([...newlog]);
        }
      }
      datasheetdata[pushindex] = singlepusheddata;
    }

    const datasheetnewrange = datasheet.getRange("A1:BL" + datasheetdata.length);
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
    const deliversheetheaderrange = deliversheet.getRange("A1:AW2");
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
    // [13]20피트 운임, [14]40피트 하이큐브 운임, [15]LCL 운임, [16]20피트 순수운임, [17]40피트 하이큐브 순수운임, [18]LCL 순수운임,
    // [19]프리타임, [20]USCAN, [21, 22, 23, 24]도착지비용1{항목, 단위, 화폐, 금액} ~ [45, 46, 47, 48]도착지비용7

    let singledata: (string | number | boolean)[] = [];

    for (let i = 2; i < datasheetdata.length; i++) {
      singledata = datasheetdata[i];
    
      // 기준운임1 설정
      let data_absoluteof1 = 0;
      let data_of1 = [
        String(singledata[7]),
        String(singledata[11]),
        String(singledata[15]),
        String(singledata[19])
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
        String(singledata[8]),
        String(singledata[12]),
        String(singledata[16]),
        String(singledata[20])
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
        String(singledata[9]),
        String(singledata[13]),
        String(singledata[17]),
        String(singledata[21])
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
        (singledata[22] as string) !== "" ? "환적" : "직항",
        singledata[22] as string,
        singledata[23] as string,
        singledata[2] as string,
        "운송 소요일",
        singledata[3] as string,
        singledata[4] as string,
        "20피트 운임", "40피트 하이큐브 운임", "LCL 운임", "20피트 순수운임", "40피트 하이큐브 순수운임", "LCL 순수운임",
        ((singledata[25] as string) !== "" && (new Date(singledata[27] as string) > new Date())) ? singledata[25] as string : "견적 시 문의",
        ((locationlist[data_podindex][1] as string) === "미국" || (locationlist[data_podindex][1] as string) === "캐나다") ? 1 : 0
      ];

      // 도착지 비용 작성
      let data_addedsurcharge = 0;
      let data_addedsurcharge_20std = 0;
      let data_addedsurcharge_40hc = 0;
      let data_addedsurcharge_lcl = 0;
      let e = 21;
      for (let ii of [28, 33, 38, 43, 48, 53, 58]) {
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
      e = 7;
      for (let ii = 0; ii < 4; ii++) {
        // 운임이 없으면 건너뜀
        if (singledata[e] as string === "" && singledata[e + 1] as string === "") {
          e += 4;
          continue;
        }
        // 출발지
        singleitem[2] = singledata[e - 1] as string;
        if (locationlist.map((row) => row[4] as string).indexOf(singleitem[2] as string) === -1) {
          console.log((i + 1) + "행 출발지 오류: " + singleitem[2] + "는(은) 유효한 출발지가 아닙니다.");
          singleitem[1] = singleitem[2];
          break;
        } else {
          singleitem[1] = locationlist[locationlist.map((row) => row[4] as string).indexOf(singleitem[2] as string)][3] as string;
        }
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
            singleitem[16] = singleitem[13];
            singleitem[13] = Math.max(
              Number(singleitem[13]) * (100 + marginsetting[0][0] as number) / 100,
              Number(singleitem[13]) + marginsetting[0][1] as number
            );
            singleitem[13] = Math.ceil(Number(singleitem[13]) / 10) * 10;
          } else if (Number(singledata[e] as string) !== 0) {
            // 운임이 절대값인 경우
            singleitem[13] = Number(singledata[e]) + data_addedsurcharge_20std;
            singleitem[16] = singleitem[13];
            singleitem[13] = Math.max(
              Number(singleitem[13]) * (100 + marginsetting[0][0] as number) / 100,
              Number(singleitem[13]) + marginsetting[0][1] as number
            );
            singleitem[13] = Math.ceil(Number(singleitem[13]) / 10) * 10;
            // 운임이 없는 경우
          } else {
            singleitem[13] = "";
            singleitem[16] = "";
          }
          // 40피트 하이큐브
          if (String(singledata[e + 1]).includes("+")) {
            // 운임이 절대값이 아닌 경우
            singleitem[14] = Number(String(singledata[e + 1]).replace("+", ""));
            singleitem[14] = Number(singleitem[14]) + data_absoluteof2 + data_addedsurcharge_40hc;
            singleitem[17] = singleitem[14];
            singleitem[14] = Math.max(
              Number(singleitem[14]) * (100 + marginsetting[1][0] as number) / 100,
              Number(singleitem[14]) + marginsetting[1][1] as number
            );
            singleitem[14] = Math.ceil(Number(singleitem[14]) / 10) * 10;
          } else if (Number(singledata[e + 1] as string) !== 0) {
            // 운임이 절대값인 경우
            singleitem[14] = Number(singledata[e + 1]) + data_addedsurcharge_40hc;
            singleitem[17] = singleitem[14];
            singleitem[14] = Math.max(
              Number(singleitem[14]) * (100 + marginsetting[1][0] as number) / 100,
              Number(singleitem[14]) + marginsetting[1][1] as number
            );
            singleitem[14] = Math.ceil(Number(singleitem[14]) / 10) * 10;
            // 운임이 없는 경우
          } else {
            singleitem[14] = "";
            singleitem[17] = "";
          }
          singleitem[15] = "";
          singleitem[18] = "";

          // LCL 운임
        } else {
          singleitem[13] = "";
          singleitem[14] = "";
          singleitem[16] = "";
          singleitem[17] = "";
          if (String(singledata[e]).includes("+")) {
            // 운임이 절대값이 아닌 경우
            singleitem[15] = Number(String(singledata[e]).replace("+", ""));
            singleitem[15] = Number(singleitem[15]) + data_absoluteof1 + data_addedsurcharge_lcl;
            singleitem[18] = singleitem[15];
            singleitem[15] = Math.max(
              Number(singleitem[15]) * (100 + marginsetting[2][0] as number) / 100,
              Number(singleitem[15]) + marginsetting[2][1] as number
            );
            singleitem[15] = Math.ceil(Number(singleitem[15]) / 10) * 10;
          } else if (Number(singledata[e] as string) !== 0) {
            // 운임이 절대값인 경우
            singleitem[15] = Number(singledata[e]) + data_addedsurcharge_lcl;
            singleitem[18] = singleitem[15];
            singleitem[15] = Math.max(
              Number(singleitem[15]) * (100 + marginsetting[2][0] as number) / 100,
              Number(singleitem[15]) + marginsetting[2][1] as number
            );
            singleitem[15] = Math.ceil(Number(singleitem[15]) / 10) * 10;
          } // LCL은 운임이 없는 경우를 생략함 (이미 조건에 포함됨)
        }

        deliversheetdata.push([...singleitem]);
        e += 4;
      }
    }

    let deliversheetrange = deliversheet.getRange("A:AW").getUsedRange();
    deliversheetrange.load("values");
    await context.sync();

    // 입력할 데이터가 이미 입력된 데이터보다 많으면 범위를 다시 지정함
    if (deliversheetdata.length > deliversheetrange.values.length) {
      deliversheetrange = deliversheet.getRange("A1:AW" + (deliversheetdata.length));
      deliversheetrange.load("values");
      await context.sync();
    }

    // 입력할 데이터가 이미 입력된 데이터보다 적으면 빈 행을 추가함
    while (deliversheetdata.length < deliversheetrange.values.length) {
      deliversheetdata.push([
        // 49개의 빈 요소 추가
        ...Array(49).fill("")
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
    const searchsettingsrange = settingsheet.getRange("C5:I9");
    const marginsettingsrange = settingsheet.getRange("C14:D16");
    const sortsettingsrange = settingsheet.getRange("D20:K20");
    searchsettingsrange.load("values");
    marginsettingsrange.load("values");
    sortsettingsrange.load("values");

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
    const sortsetting = sortsettingsrange.values[0];
    const datasheetvalues = datasheetrange.values;
    const exchangeratelist = exchangeratesheetrange.values.slice(1, undefined);
    const locationlist = locationsheetrange.values;
    const filterlist = filtersheetrange.values;
    const resultdataheader = resultdataheaderrange.values;
    let resultdata: (string | number | boolean)[][] = [];

    let searchsettings = {
      // 검색 조건
      from : searchsetting[1][0] as string,
      fromtype : (
        searchsetting[1][0] as string === "All" ? "all" :
        searchsetting[1][1] as boolean ? "filter" :
        searchsetting[1][2] as boolean ? "country" :
        searchsetting[1][3] as boolean ? "region" :
        "location"
      ),
      to : searchsetting[2][0] as string,
      totype : (
        searchsetting[2][0] as string === "All" ? "all" :
        searchsetting[2][1] as boolean ? "filter" :
        searchsetting[2][2] as boolean ? "country" :
        searchsetting[2][3] as boolean ? "region" :
        "location"
      ),
      carrier : searchsetting[3][0] as string,
      carriertype : (
        searchsetting[3][0] as string === "All" ? "all" :
        searchsetting[3][1] as boolean ? "filter" :
        "specific"
      ),
      volumetype : (
        searchsetting[4][0] as string === "LCL" ? "LCL" :
        searchsetting[4][0] as string === "All" ? "all" :
        "FCL"
      ),
      containtransshippingroute : searchsetting[4][2] as boolean,
      containexpiredfare : searchsetting[4][3] as boolean,
      // 옵션
      addmargin : searchsetting[0][4] as boolean,
      expiredfareonly : searchsetting[1][4] as boolean,
      expiredfreetimeonly : searchsetting[2][4] as boolean,
      conditionzero : searchsetting[3][4] as boolean,
      filterbyofferinglevel :
        searchsetting[4][4] as boolean === false ? "every" :
        searchsetting[4][6].substring(0, 1) === "!" ? "except" : "only",
      offeringlevel : (searchsetting[4][6] as string).replace("!", "").trim()
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
      searchsettings.fromtype === "all" ? "모두" :
        locationlist.map((row) => row[4] as string).indexOf(searchsettings.from) === -1 ? searchsettings.from :
        locationlist[locationlist.map((row) => row[4] as string).indexOf(searchsettings.from)][3],
      searchsettings.totype === "all" ? "모두" :
        locationlist.map((row) => row[4] as string).indexOf(searchsettings.to) === -1 ? searchsettings.to :
        locationlist[locationlist.map((row) => row[4] as string).indexOf(searchsettings.to)][3],
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

    for (let i = 2; i < datasheetvalues.length; i++) {
      singledata = datasheetvalues[i];
      if (
        // 출발지 확인
        (
          (searchfor.from.indexOf(String(singledata[6])) !== -1 && (String(singledata[7]) !== "" || String(singledata[8]) !== "")) ||
          (searchfor.from.indexOf(String(singledata[10])) !== -1 && (String(singledata[11]) !== "" || String(singledata[12]) !== "")) ||
          (searchfor.from.indexOf(String(singledata[14])) !== -1 && (String(singledata[15]) !== "" || String(singledata[16]) !== "")) ||
          (searchfor.from.indexOf(String(singledata[18])) !== -1 && (String(singledata[19]) !== "" || String(singledata[20]) !== "")) ||
          searchsettings.fromtype === "all"
        ) &&
        // 도착지 확인
        (searchfor.to.indexOf(singledata[0] as string) !== -1 || searchsettings.totype === "all" ) &&
        // 운송사 확인
        (searchfor.carrier.indexOf(singledata[1] as string) !== -1 || searchsettings.carriertype === "all") &&
        // 운송 단위 확인
        (singledata[2] as string === searchsettings.volumetype || searchsettings.volumetype === "all") &&
        // 환적 경로 확인
        (singledata[22] as string === "" || searchsettings.containtransshippingroute) &&
        // 만료 데이터 확인
        (new Date(singledata[4] as string) >= new Date() || searchsettings.containexpiredfare || searchsettings.expiredfareonly) &&
        // 운임 만료 데이터만 표시 옵션일 경우 만료 데이터만 포함
        (!searchsettings.expiredfareonly || new Date(singledata[4] as string) < new Date()) &&
        // 프리타임 만료 데이터만 표시 옵션일 경우 만료 데이터만 포함
        (!searchsettings.expiredfreetimeonly || new Date(singledata[27] as string) < new Date()) &&
        // 운임 특성 필터
        (searchsettings.filterbyofferinglevel === "every" ||
          (searchsettings.filterbyofferinglevel === "only" && String(singledata[5]).includes(searchsettings.offeringlevel)) ||
          (searchsettings.filterbyofferinglevel === "except" && !String(singledata[5]).includes(searchsettings.offeringlevel))
        )
      ) {
        // 기준운임1 설정
        let data_absoluteof1 = 0;
        let data_of1 = [
          String(singledata[7]),
          String(singledata[11]),
          String(singledata[15]),
          String(singledata[19])
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
          String(singledata[8]),
          String(singledata[12]),
          String(singledata[16]),
          String(singledata[20])
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
          String(singledata[9]),
          String(singledata[13]),
          String(singledata[17]),
          String(singledata[21])
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
        for (let ii of [28, 33, 38, 43, 48, 53, 58]) {
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
        searchresult[1] = singledata[22];
        searchresult[2] = singledata[23];
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

        // 선적항이 출발지 검색 조건에 포함되고 운임이 존재하는 경우 해당 선적항 정보를 검색 결과에 추가 (출발 위치, 소요일, 운임)
        for (let e of [6, 10, 14, 18]) {
          if (
            (searchfor.from.indexOf(String(singledata[e])) !== -1 || searchsettings.fromtype === "all") &&
            (String(singledata[e + 1]) !== "" || String(singledata[e + 2]) !== "")
          ) {
            // 출발 위치
            searchresult[0] = singledata[e] as string;
            if (locationlist.map((row) => row[4] as string).indexOf(searchresult[0] as string) === -1) {
              console.log((i + 1) + "행 출발지 오류: " + searchresult[0] + "는(은) 유효한 출발지가 아닙니다.");
              searchresult[0] = searchresult[0];
            } else {
              searchresult[0] = locationlist[locationlist.map((row) => row[4] as string).indexOf(searchresult[0] as string)][3] as string;
            }
            // 운송 소요일
            let data_tt_formatted = "";
            // 운송 소요일이 절대값이 아닌 경우
            if (String(singledata[e + 3]).includes("+")) {
              // 기준 소요일이 범위로 지정된 경우
              if (String(data_absolutett).includes("~")) {
                let data_tt_range = data_absolutett.split("~").map((item) => item.trim());
                data_tt_formatted = (
                  (Number(data_tt_range[0]) + Number(String(singledata[e + 3]).replace("+", ""))) +
                  "~" +
                  (Number(data_tt_range[1]) + Number(String(singledata[e + 3]).replace("+", "")))
                );
              // 기준 소요일이 범위로 지정되지 않은 경우
              } else {
                data_tt_formatted = (
                  String(Number(data_absolutett) + Number(String(singledata[e + 3]).replace("+", "")))
                );
              }
              // 운송 소요일이 절대값인 경우
            } else if ((singledata[e + 3] as string) !== "") {
              data_tt_formatted = singledata[e + 3] as string;
              // 운송 소요일이 없는 경우
            } else {
              data_tt_formatted = "견적 시 문의";
            }
            searchresult[4] = data_tt_formatted;
            
            // 운임
            let data_fare_formatted = "";
            let data_fare = 0;
            searchresult[8] = 0;
              // FCL 운임
            if (singledata[2] === "FCL") {
              // 20피트
              if (String(singledata[e + 1]).includes("+")) {
                // 운임이 절대값이 아닌 경우
                data_fare = Number(String(singledata[e + 1]).replace("+", ""));
                data_fare = Number(data_fare) + data_absoluteof1 + data_addedsurcharge_20std;
                data_fare = Math.max(
                  Number(data_fare) * (100 + marginsetting[0][0] as number) / 100,
                  Number(data_fare) + marginsetting[0][1] as number
                );
                data_fare = Math.ceil(Number(data_fare) / 10) * 10;
                data_fare_formatted = String(data_fare);
              } else if (Number(singledata[e + 1] as string) !== 0) {
                // 운임이 절대값인 경우
                data_fare = Number(singledata[e + 1]) + data_addedsurcharge_20std;
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
              if (sortsetting[7] === true || (sortsetting[7] === false && data_fare === 0)) {
                searchresult[8] = 31416; // 임시로 매우 큰 수를 넣음 (운임이 없는 경우 맨 뒤로 정렬하기 위함)
              } else {
                searchresult[8] = data_fare; // 20피트 운임으로 정렬하도록 설정했고 20피트 운임이 있다면 이 운임을 임시로 저장
              }

              data_fare_formatted = data_fare_formatted + " | ";

              // 40피트 하이큐브
              if (String(singledata[e + 2]).includes("+")) {
                // 운임이 절대값이 아닌 경우
                data_fare = Number(String(singledata[e + 2]).replace("+", ""));
                data_fare = Number(data_fare) + data_absoluteof2 + data_addedsurcharge_40hc;
                data_fare = Math.max(
                  Number(data_fare) * (100 + marginsetting[1][0] as number) / 100,
                  Number(data_fare) + marginsetting[1][1] as number
                );
                data_fare = Math.ceil(Number(data_fare) / 10) * 10;
                data_fare_formatted = data_fare_formatted + String(data_fare);
              } else if (Number(singledata[e + 2] as string) !== 0) {
                // 운임이 절대값인 경우
                data_fare = Number(singledata[e + 2]) + data_addedsurcharge_40hc;
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
              if (sortsetting[7] === false || (sortsetting[7] === true && data_fare === 0)) {
                searchresult[8] = 31416; // 임시로 매우 큰 수를 넣음 (운임이 없는 경우 맨 뒤로 정렬하기 위함)
              } else {
                searchresult[8] = data_fare; // 40피트 하이큐브 운임으로 정렬하도록 설정했고 40피트 하이큐브 운임이 있다면 이 운임을 임시로 저장
              }

              // LCL 운임
            } else {
              if (String(singledata[e + 1]).includes("+")) {
                // 운임이 절대값이 아닌 경우
                data_fare = Number(String(singledata[e + 1]).replace("+", ""));
                data_fare = Number(data_fare) + data_absoluteof1 + data_addedsurcharge_lcl;
                data_fare = Math.max(
                  Number(data_fare) * (100 + marginsetting[2][0] as number) / 100,
                  Number(data_fare) + marginsetting[2][1] as number
                );
                data_fare = Math.ceil(Number(data_fare) / 10) * 10;
                data_fare_formatted = String(data_fare);
              } else if (Number(singledata[e + 1] as string) !== 0) {
                // 운임이 절대값인 경우
                data_fare = Number(singledata[e + 1]) + data_addedsurcharge_lcl;
                data_fare = Math.max(
                  Number(data_fare) * (100 + marginsetting[2][0] as number) / 100,
                  Number(data_fare) + marginsetting[2][1] as number
                );
                data_fare = Math.ceil(Number(data_fare) / 10) * 10;
                data_fare_formatted = String(data_fare);
              }
              if (searchsettings.volumetype === "all") {
                data_fare_formatted = data_fare_formatted + " (LCL)";
              }
              /*  // 정렬
               *  LCL 운임이 검색 결과에 포함되는건
               *  1. 운송 단위 조건이 LCL인 경우
               *  2. 운송 단위 조건이 ALL이고 LCL 운임이 존재하는 경우
               *  뿐인데 1번이면 그냥 LCL 운임으로 정렬될 것이고, 2번이면 LCL 운임이 모든 컨테이너 운임보다 낮기 때문에 LCL이 맨 앞에 올 것임
               */ 
              searchresult[8] = data_fare;
            }
            searchresult[6] = data_fare_formatted;
            resultdata.push([...searchresult]);
          } // for(let e of [6, 10, 14, 18])
        } // if(선적항이 출발지 검색 조건에 포함되고 운임이 존재하는 경우)
      } // if(검색 조건에 맞는 경우)
    } // for(let i = 2; i < datasheetvalues.length; i++)

    // 정렬
    if (sortsetting[6] === true) { // 오름차순
      resultdata.sort((a, b) => {return (a[8] as number) - (b[8] as number);});
    } else { // 내림차순
      resultdata.sort((a, b) => {return (b[8] as number) - (a[8] as number);});
    }
    // 그룹화
    let groupingorder: (number | string)[][] = [
      ["출발 위치", sortsetting[0]], ["운송사", sortsetting[2]], ["도착 위치", sortsetting[4]]];
    groupingorder = groupingorder.filter((item) => item[1] !== 0);
    groupingorder = groupingorder.sort((a, b) => (b[1] as number) - (a[1] as number));
    let listoforigins: string[] = [];
    let listofcarriers: string[] = [];
    let listofdestinations: string[] = [];
    if (groupingorder.length > 0) {
      for (let i = 0; i < resultdata.length; i++) {
        if (listoforigins.indexOf(resultdata[i][0] as string) === -1) {
          listoforigins.push(resultdata[i][0] as string);
        }
        if (listofcarriers.indexOf(resultdata[i][5] as string) === -1) {
          listofcarriers.push(resultdata[i][5] as string);
        }
        if (listofdestinations.indexOf(resultdata[i][3] as string) === -1) {
          listofdestinations.push(resultdata[i][3] as string);
        }
      }
      for (let groupinginstance of groupingorder) {
        if (groupinginstance[0] === "출발 위치") {
          for (let i = 0; i < resultdata.length; i++) {
            resultdata[i][8] = listoforigins.indexOf(resultdata[i][0] as string);
            resultdata[i][8] = Number(resultdata[i][8]) * 31416 + i;
          }
        } else if (groupinginstance[0] === "운송사") {
          for (let i = 0; i < resultdata.length; i++) {
            resultdata[i][8] = listofcarriers.indexOf(resultdata[i][5] as string);
            resultdata[i][8] = Number(resultdata[i][8]) * 31416 + i;
          }
        } else if (groupinginstance[0] === "도착 위치") {
          for (let i = 0; i < resultdata.length; i++) {
            resultdata[i][8] = listofdestinations.indexOf(resultdata[i][3] as string);
            resultdata[i][8] = Number(resultdata[i][8]) * 31416 + i;
          }
        }
        resultdata.sort((a, b) => {return (a[8] as number) - (b[8] as number);});
      }
    }
    // 정렬 및 그룹화 완료
    for (let i = 0; i < resultdata.length; i++) {
      resultdata[i].splice(8, 1);
    }

    resultdata.unshift(resultdataheader[1]);
    resultdata.unshift(resultdataheader[0]);

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
    // 전체, 유효 기간, 운임 특성, 선적항1 정보, 선적항2 정보, 선적항3 정보, 선적항4 정보, 경로 정보, 메모, 프리타임, 도착지 비용1~7 (총 17개)

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
        datasheetdata[dataindex] = ([singlelog[3], singlelog[4], singlelog[5]] as string[]).concat(Array(61).fill(""));
      } else if (singlelog[2] === "수정") {
        if (singlelog[6] === "유효 기간") {
          datasheetdata[dataindex][3] = singlelog[7];
          datasheetdata[dataindex][4] = singlelog[8];
        } else if (singlelog[6] === "운임 특성") {
          datasheetdata[dataindex][5] = singlelog[7];
        } else if (singlelog[6] === "선적항1 정보") {
          datasheetdata[dataindex][6] = singlelog[7];
          datasheetdata[dataindex][7] = singlelog[8];
          datasheetdata[dataindex][8] = singlelog[9];
          datasheetdata[dataindex][9] = singlelog[10];
        } else if (singlelog[6] === "선적항2 정보") {
          datasheetdata[dataindex][10] = singlelog[7];
          datasheetdata[dataindex][11] = singlelog[8];
          datasheetdata[dataindex][12] = singlelog[9];
          datasheetdata[dataindex][13] = singlelog[10];
        } else if (singlelog[6] === "선적항3 정보") {
          datasheetdata[dataindex][14] = singlelog[7];
          datasheetdata[dataindex][15] = singlelog[8];
          datasheetdata[dataindex][16] = singlelog[9];
          datasheetdata[dataindex][17] = singlelog[10];
        } else if (singlelog[6] === "선적항4 정보") {
          datasheetdata[dataindex][18] = singlelog[7];
          datasheetdata[dataindex][19] = singlelog[8];
          datasheetdata[dataindex][20] = singlelog[9];
          datasheetdata[dataindex][21] = singlelog[10];
        } else if (singlelog[6] === "경로 정보") {
          datasheetdata[dataindex][22] = singlelog[7];
          datasheetdata[dataindex][23] = singlelog[8];
          datasheetdata[dataindex][24] = singlelog[9];
        } else if (singlelog[6] === "프리타임") {
          datasheetdata[dataindex][25] = singlelog[7];
          datasheetdata[dataindex][26] = singlelog[8];
          datasheetdata[dataindex][27] = singlelog[9];
        } else if (singlelog[6] === "도착지 비용1") {
          datasheetdata[dataindex][28] = singlelog[7] === "true";
          datasheetdata[dataindex][29] = singlelog[8];
          datasheetdata[dataindex][30] = singlelog[9];
          datasheetdata[dataindex][31] = singlelog[10];
          datasheetdata[dataindex][32] = Number(singlelog[11]);
        } else if (singlelog[6] === "도착지 비용2") {
          datasheetdata[dataindex][33] = singlelog[7] === "true";
          datasheetdata[dataindex][34] = singlelog[8];
          datasheetdata[dataindex][35] = singlelog[9];
          datasheetdata[dataindex][36] = singlelog[10];
          datasheetdata[dataindex][37] = Number(singlelog[11]);
        } else if (singlelog[6] === "도착지 비용3") {
          datasheetdata[dataindex][38] = singlelog[7] === "true";
          datasheetdata[dataindex][39] = singlelog[8];
          datasheetdata[dataindex][40] = singlelog[9];
          datasheetdata[dataindex][41] = singlelog[10];
          datasheetdata[dataindex][42] = Number(singlelog[11]);
        } else if (singlelog[6] === "도착지 비용4") {
          datasheetdata[dataindex][43] = singlelog[7] === "true";
          datasheetdata[dataindex][44] = singlelog[8];
          datasheetdata[dataindex][45] = singlelog[9];
          datasheetdata[dataindex][46] = singlelog[10];
          datasheetdata[dataindex][47] = Number(singlelog[11]);
        } else if (singlelog[6] === "도착지 비용5") {
          datasheetdata[dataindex][48] = singlelog[7] === "true";
          datasheetdata[dataindex][49] = singlelog[8];
          datasheetdata[dataindex][50] = singlelog[9];
          datasheetdata[dataindex][51] = singlelog[10];
          datasheetdata[dataindex][52] = Number(singlelog[11]);
        } else if (singlelog[6] === "도착지 비용6") {
          datasheetdata[dataindex][53] = singlelog[7] === "true";
          datasheetdata[dataindex][54] = singlelog[8];
          datasheetdata[dataindex][55] = singlelog[9];
          datasheetdata[dataindex][56] = singlelog[10];
          datasheetdata[dataindex][57] = Number(singlelog[11]);
        } else if (singlelog[6] === "도착지 비용7") {
          datasheetdata[dataindex][58] = singlelog[7] === "true";
          datasheetdata[dataindex][59] = singlelog[8];
          datasheetdata[dataindex][60] = singlelog[9];
          datasheetdata[dataindex][61] = singlelog[10];
          datasheetdata[dataindex][62] = Number(singlelog[11]);
        } else if (singlelog[6] === "메모") {
          datasheetdata[dataindex][63] = singlelog[7];
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
async function tryCatch(callback: () => Promise<void>) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}
