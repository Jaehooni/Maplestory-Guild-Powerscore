const baseUrl = 'https://open.api.nexon.com';
// 본인의 API 키 집어넣으면 됨
const API_KEY = ""
const params = {
  'method': 'get',
  'contentType': 'application/json',
  'headers': {
    'x-nxopen-api-key': API_KEY,
    'accept': 'application/json'
  }
}

function main(){
  // Sheet에서 Button을 클릭하여 Script를 수행하기 위한 코드
  let ui = SpreadsheetApp.getUi();
  let alert = ui.alert("경고", "불러오는 데 약 2분 정도가 소요됩니다. 정말 불러오시겠습니까?", ui.ButtonSet.YES_NO);
  let result;
  // result = getInfo();

  if (alert == ui.Button.YES) {
      try{
        result = getInfo();
      } catch (err){
        ui.alert("정보를 불러오는데 실패하였습니다.");
      }
  }

  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("길드원 명단");

  for(let i=3; i<=202; i++) {
    sheet.getRange(i, 2).setValue("");
    sheet.getRange(i, 3).setValue("");
    sheet.getRange(i, 4).setValue("");
    sheet.getRange(i, 5).setValue("");
  }

  for(let i=0; i< result.members.length; i++) {
    sheet.getRange(i+3, 2).setValue(result.members[i]);
    Logger.log(result.members[i]);
    if (result.memberinfos[i] === false){
      sheet.getRange(i+3, 3).setValue("");
      sheet.getRange(i+3, 4).setValue("");
      sheet.getRange(i+3, 5).setValue("");
    }

    else{
      sheet.getRange(i+3, 3).setValue(result.memberinfos[i]["class"]);
      sheet.getRange(i+3, 4).setValue(result.memberinfos[i]["level"]);
      sheet.getRange(i+3, 5).setValue(result.memberinfos[i]["power"]);
    }
  }

  sheet.getRange(4, 7).setValue(result.members.length);
  sheet.getRange(4, 8).setValue(getDateInfo());


}


function getInfo(){
  const ocidList = [];
  const memberInfoList = [];
  // const powerList = [];
  const memberList = getGuildMemberList();
  memberList.forEach(function(member){
    try{
    ocidList.push(getGuildMemberOCID(member));
    } catch (err){
      ocidList.push(false);
    }
  })

  ocidList.forEach(function(ocid){
    if (ocid === false){
      memberInfoList.push(false);
    }

    else{
      const info = getGuildMemberInfo(ocid);
      const power = getGuildMemberPower(ocid);
      memberInfoList.push(
        {
          "class": info[0],
          "level": info[1],
          "power": power
        }
      )
    }
  })

  const result = {
    'members': memberList,
    'memberinfos': memberInfoList
  }

  // console.log(memberList.length);
  // console.log(memberList);
  // console.log(memberInfoList.length);
  // console.log(memberInfoList);

  return result;
}

function makeUrl(baseurl, context){
  return baseurl + context;
}

function getDateInfo(){
  let date = new Date();
  date.setDate(date.getDate() - 1);
  const year = date.getFullYear();
  const month = ('0' + (date.getMonth() + 1)).slice(-2);
  const day = ('0' + date.getDate()).slice(-2);

  const dateParam = `${year}-${month}-${day}`
  return dateParam
}

// KLAS 이외의 길드 사용 시 조정해서 사용하면 됨
function getGuid() {
  const guild_name = 'KLAS';
  const world_name = '크로아';

  const context = `/maplestory/v1/guild/id?guild_name=${guild_name}&world_name=${world_name}`;
  const url = makeUrl(baseUrl, context);
  //console.log(url);

  const response = UrlFetchApp.fetch(url, params);
  const jsondata = JSON.parse(response.getContentText());
  return jsondata.oguild_id;
}

function getGuildMemberList(){
  const date = getDateInfo();
  const guid = getGuid();
  const context = `/maplestory/v1/guild/basic?oguild_id=${guid}&date=${date}`
  const url = makeUrl(baseUrl, context);
  //console.log(url);

  const response = UrlFetchApp.fetch(url, params);
  const jsondata = JSON.parse(response.getContentText());
  return jsondata.guild_member;
}

function getGuildMemberOCID(memberName){
  const context = `/maplestory/v1/id?character_name=${memberName}`;
  const url = makeUrl(baseUrl, context);

  const response = UrlFetchApp.fetch(url, params);
  const jsondata = JSON.parse(response.getContentText());

  return jsondata.ocid;
}

function getGuildMemberInfo(memberOcid){
  const date = getDateInfo();
  const context = `/maplestory/v1/character/basic?ocid=${memberOcid}&date=${date}`
  const url = makeUrl(baseUrl, context);

  const response = UrlFetchApp.fetch(url, params);
  const jsondata = JSON.parse(response.getContentText());

  return [jsondata.character_class, jsondata.character_level];
}

function getGuildMemberPower(memberOcid){
  const date = getDateInfo();
  const context = `/maplestory/v1/character/stat?ocid=${memberOcid}&date=${date}`
  const url = makeUrl(baseUrl, context);

  const response = UrlFetchApp.fetch(url, params);
  const jsondata = JSON.parse(response.getContentText());

  //멤버의 전투력만 가져 옴
  return jsondata.final_stat[42].stat_value;
}
