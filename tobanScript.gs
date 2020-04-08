var cal = CalendarApp.getCalendarById('');
var sheet = SpreadsheetApp.getActiveSheet();

var options =
    {
      description: "電話当番"
    }

var nameArray = [];
var groupArray = [];
var restNameArray = [];
var penaltyNameArray = [];

var year = "";
var month = "";
var first_day = "";
var last_day = "";

var nx = 2;
var numOfShifts = 2;
var penaltyNameX = 4;

var finalDecisionList = [];

/*
 * エントリーポイント
 */
function main(){
  readPenaltyMember(); // ペナルティーメンバーを読む
  Logger.log(penaltyNameArray);

  readSheet(); // メンバーを読む
  readData();  // 日付を読む
  Logger.log(nameArray);

  selectMembers();
  Logger.log(finalDecisionList);

  resetNumOfShifts();

  if (numOfShifts == 2) {
    setCalender２persons();
  } else if (numOfShifts == 1) {
    setCalender();
  }

}

// 期間やシフトの人数などのデータを読む
function readData() {
  year = sheet.getRange(2, 7).getDisplayValue().toString();
  month = sheet.getRange(2, 8).getDisplayValue().toString();
  first_day = sheet.getRange(2, 9).getDisplayValue().toString();
  last_day = sheet.getRange(3, 9).getDisplayValue().toString();
  numOfShifts = sheet.getRange(4, 9).getDisplayValue().toString();
  Logger.log(year + '年' + month + '月' + first_day + '日 ~ '+ last_day + '日');
}

/*
 * ペナルティ者の名前をpenaltyNameArrayに追加する
 */
function readPenaltyMember() {
  for(var i = 2; i <= sheet.getLastRow(); i++) {
    var name = sheet.getRange(i, 4).getDisplayValue().toString();
    var penalty = sheet.getRange(i, 5).getDisplayValue().toString();

    // nameが書いてない場合continue
    if (name.length <= 0) {
      continue;
    }

    for(var j = 0; j < penalty; j++) {
       penaltyNameArray.push(name);
       reducePenaltyNum(penaltyNameX, i, sheet);
    }
  }
}

// ゼロかどうか
function isZero(pnx) {
  if (pnx == 0) {
    return true;
  }
  return false;
}

/*
 * （登録したら）ペナルティーの数を減らす
 */
function reducePenaltyNum(x, y) {
  var val = sheet.getRange(y, x).getValue();
  if (isZero(val)) {
    return;
  }
  sheet.getRange(y, x+1).setValue(0);
}

// シートを読む
function readSheet() {
  var MinNum = getMinNumOfShifts(sheet);
  for(var i = 2; i <= sheet.getLastRow(); i++) {
    var name = sheet.getRange(i, 1).getDisplayValue().toString();
    var nowNum = sheet.getRange(i, 2).getDisplayValue().toString();
    var active = sheet.getRange(i, 3).getDisplayValue().toString();

    if (active == "n") {
      continue;
    }

    if (nowNum == MinNum) {
      restNameArray.push(name);
      setNum(nx, i);
    }
    nameArray.push(name)
  }
}


// メンバーを選択する
function selectMembers() {
  allMemberN = allMemberNum();
  var selected_members;

  // 初回はペナルティのメンバーを含む
  var arr = penaltyNameArray.concat(restNameArray);
  allMemberN -= arr.length;
  Logger.log('初回 ' + arr);
  finalDecisionList = finalDecisionList.concat(shuffle(arr));


  while(allMemberN > nameArray.length) {
    var nameList = nameArray.slice();
    finalDecisionList = finalDecisionList.concat(shuffle(nameList)); // ペナルティー以外のメンバー
    allMemberN -= nameArray.length;
  }

  finalDecisionList = finalDecisionList.concat(shuffle(nameArray.slice(0, allMemberN)));
  addNumOfShifts(0, allMemberN);
}

/*
 * シフトをこなしたらシートのやった回数を追加する
 */
function addNumOfShifts(startMemberN, lastMemberN) {
  for(var i = startMemberN+2; i < lastMemberN+2; i++) {
    var nowNum = sheet.getRange(i, 2).getDisplayValue().toString();
    var active = sheet.getRange(i, 3).getDisplayValue().toString();

    if (active == "n") {
      lastMemberN += 1
      continue;
    }

    setNum(nx, i);
  }
}

function setNum(nx, ny) {
  var val = sheet.getRange(ny, nx).getValue();
  sheet.getRange(ny, nx).setValue(val+1);
}

// 必要となるメンバーの数を取得する
function allMemberNum() {
  return (Number(last_day) - Number(first_day)+1) * numOfShifts; // 日数*2
}

// 配列をシャッフルする
function shuffle(nameList) {
  Logger.log('シャッフル前 ' + nameList);
  var members = [];
  limit = limitNum(nameList);
  for (var i = 0; i < limit; i++) {
    var arrayIndex = Math.floor(Math.random() * nameList.length);
    members[i] = nameList[arrayIndex];
    // 1回選択された値は削除して再度選ばれないようにする
    nameList.splice(arrayIndex, 1);
  }
  Logger.log('シャッフル後 ' + members);
  return members;
}

function limitNum(nameList) {
  var minN;
  if (nameList.length <= allMemberNum()) {
    minN = nameList.length;
  } else {
    minN = allMemberNum();
  }
  Logger.log('人数　' + minN);
  return minN;
}

function getMinNumOfShifts() {
  var MinNum = 1000000;
  for(var i = 2; i <= sheet.getLastRow(); i++) {
    var nowNum = sheet.getRange(i, nx).getDisplayValue().toString();
    var active = sheet.getRange(i, 3).getDisplayValue().toString();

    if (active == "n") {
      continue;
    }

    if (MinNum > Number(nowNum)) {
      MinNum = Number(nowNum);
    }
  }
  return MinNum;
}

/*
 * スプレッドシートのシフトの回数をリセットする
 */
function resetNumOfShifts() {
  var MinNum = getMinNumOfShifts(sheet);
  for(var i = 2; i <= sheet.getLastRow(); i++) {
    var nowNum = sheet.getRange(i, nx).getDisplayValue().toString();
    var active = sheet.getRange(i, 3).getDisplayValue().toString();

    sheet.getRange(i, nx).setValue(Number(nowNum) - Number(MinNum));

    if (active == "n") {
      sheet.getRange(i, nx).setValue(0);
    }
  }
}

function setCalender２persons() {
  var groupArray = [];
  for(var i=0; i<allMemberNum()-1; i+=2) {
    var name1 = finalDecisionList[i];
    var name2 = finalDecisionList[i+1];
    groupArray.push(name1 + "・" + name2);
  }

  for(var i=0; i<groupArray.length; i++){
    var n = groupArray[i];
    var d = (i+Number(first_day)).toString();
    var mon = ( '00' + month ).slice( -2 );
    var startTime = new Date("2020/" + mon + "/" + d + " 18:00");
    var endTime = new Date("2020/" + mon + "/" + d + " 21:00");

    Logger.log((i+1).toString(), n)
    // cal.createEvent(n, startTime, endTime, options);
  }
}

function setCalender() {
  var groupArray = [];
  for(var i=0; i<allMemberNum(); i++) {
    var n = finalDecisionList[i];
    var d = (i+Number(first_day)).toString();
    var mon = ( '00' + month ).slice( -2 );
    var startTime = new Date("2020/" + mon + "/" + d + " 18:00");
    var endTime = new Date("2020/" + mon + "/" + d + " 21:00");

    Logger.log((i+1).toString(), n)
  }
}