const config = require('config');
const axios = require("axios");
const cheerio = require("cheerio");
const fs = require('fs').promises;
const path = require('path');
const {authenticate} = require('@google-cloud/local-auth');
const {google} = require('googleapis');
const log4js = require("log4js");

//是否是在Debug
const isDebug = config.get("Debug").toLowerCase() === "true";




//公開資訊觀測站法說會查詢網址
const mopsConferenceUrl = "https://mops.twse.com.tw/mops/web/t100sb07_1";
// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/calendar.readonly','https://www.googleapis.com/auth/calendar'];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = path.join(process.cwd(), 'token.json');
const CREDENTIALS_PATH = path.join(process.cwd(), 'credentials.json');
// 獲取當前日期
const currentDate = formatDate(new Date());
//log Level
const logLevel = config.get("LogLevel");

const stockNumbers= (config.get("StockNumbers")).split(",");


//#region class
//法說會資訊
class Conference {
    constructor(data){
        this.CompId = data.CompId;
        this.CompName = data.CompName;
        this.Time = data.Time;
        this.Location = data.Location;
        this.Content = data.Content;
    };
}

//重訊
class MaterialInfo{
    constructor(data){
        this.CompId = data.CompId;
        this.CompName = data.CompName;
        this.SpeakDate = data.SpeakDate;
        this.SpeakTime = data.SpeakTime;
        this.Subject = data.Subject;
        this.FactDate = data.FactDate;
        this.Content = data.Content;
    }
}

//google行事曆 event
class EventDateTime {
    constructor(data){
        this.dateTime = data.dateTime;
        this.timeZone = data.timeZone ?? "Asia/Taipei";
    }
}
  
class EventAttendee {
    constructor(data){
        this.email = data.email;
    }
}
  
class EventReminder {
    constructor(data){
        this.method = data.method;
        this.minutes = data.minutes;
    }
}
  
class EventReminders {
    constructor(data){
        this.useDefault = data.useDefault;
        this.overrides = data.overrides.map(item => new EventReminder(item));
    }
}
  
class CalendarEvent {
    constructor(data){
        this.summary = data.summary;
        this.location = data.location;
        this.description = data.description;
        this.start = new EventDateTime(data.start);
        this.end = new EventDateTime(data.end);
        this.recurrence = data.recurrence;
        this.attendees = data.attendees.map(item => new EventAttendee(item));
        this.reminders = new EventReminders(data.reminders);
    }
}

//#endregion

//#region log4j config
log4js.configure({
  appenders: { app: { type: "file", filename: `Log/stock-info-collector-${currentDate}.log` } },
  categories: { default: { appenders: ["app"], level: logLevel } },
});
const logger = log4js.getLogger();
//#endregion

startApp();
// conferenceCrawler();
  // TestInsertEvent("2024-05-14T07:28:43.055Z")


/**
 * 開始執行程式
 *
 */
async function startApp(){
  logger.info("===開始進行法說會爬蟲===");
  logger.info("本次查詢股號如下:",stockNumbers);
  console.log("stockNumbers:",stockNumbers);

  try{
    await getConferenceThenInsertEvent();
  }catch(ex){
    console.error('法說會爬蟲發生錯誤：' + ex);
    logger.error('法說會爬蟲發生錯誤：' + ex);
    //需要等logger紀錄之後再exit
    await delay("5000");
    return process.exit(1);
  }
  
  
  logger.info("===法說會爬蟲結束===");
}

/**
 * 爬完法說會後進行行事曆新增事件
 *
 */
async function getConferenceThenInsertEvent(){
  try{
    let conferences = await conferenceCrawler();
    //now
    let nowDate = new Date()
    console.log("===getConferenceThenInsertEvent===");
    console.table(conferences);
    const dateTimePattern = /(?<ydm>\d*\/\d*\/\d*)(?:\D|\d)*?時間：(?<hour>\d*?)點(?<min>\d*?)分/;
    for(let conference of conferences){

        let compId = conference.CompId;
        let compName = conference.CompName;
        let summary = `${compId}-${compName} 法說會`
        logger.debug(`[getConferenceThenInsertEvent]當前處理${summary}`);

        //需要再處理 目前=> 113/04/18時間：\n14 點 0 分 (24小時制)
        //把時間字串用正則取出，並轉換成ISO format: YYYY-MM-DDThh:mm:ss
        let timeStr = conference.Time?.replaceAll("\n","")?.replaceAll(" ","");
        let match =  dateTimePattern.exec(timeStr);
        let date = match.groups.ydm;
        let dateArr = date.split("/");
        let twYear = dateArr[0];
        //民國年轉西元年
        let year = Number(twYear)+1911;
        let month = dateArr[1];
        let day = dateArr[2];
        let hour = match.groups.hour;
        let min = match.groups.min;
        //month 是用monthIndex 實際月份要減1才能取到對到index
        let startDateTime = new Date(year,Number(month)-1, Number(day),  Number(hour), Number(min));
        //結束時間先加2小時
        let endDateTime = new Date(year,Number(month)-1, Number(day),  Number(hour)+2, Number(min));

        //如果小於當前時間，代表已經過去不用加入
        if(startDateTime < nowDate){
          continue;
        }
        
        logger.debug(`[getConferenceThenInsertEvent]進行到checkSameEventIsExist`);    
        let isEventExist = await checkSameEventIsExist("primary", summary, startDateTime, endDateTime);
        //如果已經再行事曆存在，也不用加入
        if(isEventExist){
          continue;
        }
        logger.debug(`[getConferenceThenInsertEvent]進行到checkSameEventIsExist`);   

        await InsertEvents(new CalendarEvent({
            summary: summary,
            location: conference.Location,
            description: conference.Content,
            start: {
                dateTime: startDateTime.toISOString(),
                timeZone: 'Asia/Taipei'
            },
            end: {
                dateTime: endDateTime.toISOString(),
                timeZone: 'Asia/Taipei'
            },
            recurrence: [],
            attendees: [],
            reminders: {
                useDefault: false,
                overrides: [
                    { method: 'email', minutes: 24 * 60 },
                    { method: 'popup', minutes: 30 }
                ]
            }
        }));
    }
  }catch(ex){    
    console.error('[getConferenceThenInsertEvent]' + ex);
    logger.error('[getConferenceThenInsertEvent Error] ' + ex);
    throw ex;
  }
}

/**
 * 日期格式化
 * @returns{string}
 */ 
function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}${month}${day}`;
}

//#region 爬蟲

/**
 * 法說會爬蟲
 *
 * @return {Promise<Conference[]|null>}
 */
async function conferenceCrawler() {
    let conferenceArr = [];
    const pattern = /<b>公司代號：<\/b>(?<compId>\d+).*?<b>公司名稱：<\/b>(?<compName>.+?)<br>/;

    //for loop 把股號拿去查詢
    for(let stockNumber of stockNumbers){
        console.log(`In conferenceCrawler loop now stock is ${stockNumber}`)
        logger.debug(`[conferenceCrawler]In conferenceCrawler loop now stock is ${stockNumber}`);

        //20250301 公開觀測資訊站網站改版，需要先查詢後再用查到的Url 做爬蟲
        var queryResponse = await fetch("https://mops.twse.com.tw/mops/api/redirectToOld", {
          "headers": {
            "accept": "*/*",
            "accept-language": "zh-TW,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "content-type": "application/json",
            "sec-ch-ua": "\"Not(A:Brand\";v=\"99\", \"Microsoft Edge\";v=\"133\", \"Chromium\";v=\"133\"",
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": "\"Windows\"",
            "sec-fetch-dest": "empty",
            "sec-fetch-mode": "cors",
            "sec-fetch-site": "same-origin",
            "cookie": "_gid=GA1.3.780466755.1740832872; _ga_WS2SDNL5XZ=GS1.3.1740832871.1.0.1740832871.0.0.0; _ga=GA1.1.1553753193.1740832872; _ga_LTMT28749H=GS1.1.1740833053.1.1.1740833619.0.0.0"
          },
          "referrerPolicy": "no-referrer",
          "body": "{\"apiName\":\"ajax_t100sb07_1\",\"parameters\":{\"co_id\":\"2330\",\"encodeURIComponent\":1,\"step\":1,\"firstin\":1,\"off\":1,\"TYPEK\":\"all\"}}",
          "method": "POST"
        });
        
        let result = await queryResponse.json();
         
        let conferenceUrl = result.result.url;

        let res = await axios.get(conferenceUrl);
        let data = res.data,

        $ = cheerio.load(data);
        let contentHtml = $("center").html()?.replace(/\s+/g, '');
        // console.log(`${stockNumber} contentHtml:`,contentHtml);
        let match =  pattern.exec(contentHtml);
        // console.log("mat:",match);
        if (match && match.groups) {     
            console.log("match!!")       
            let compId = match.groups.compId;
            let compName = match.groups.compName;
            let time =  $("center > form > table > tbody > tr:nth-child(1) > td:nth-child(3)").text();
            let location = $("center > form > table > tbody > tr:nth-child(2) > td:nth-child(2)").text();
            let content = $("center > form > table > tbody > tr:nth-child(3) > td:nth-child(2)").text();
            console.log(`compId: ${compId}`);
            console.log(`compName: ${compName}`);
            conferenceArr.push(new Conference({CompId: compId, CompName: compName, Time: time, Location: location, Content: content}));
        }
        
        console.log(`====== next comp ======`)
    }    
    // console.table(conferenceArr);
    logger.info('[conferenceCrawler]已查詢出以下資訊: ' +  JSON.stringify(conferenceArr));
    return conferenceArr;
};


function delay(time) {
    return new Promise(function(resolve) { 
        setTimeout(resolve, time)
    });
 }

//#endregion

//#region google api
/**
 * Reads previously authorized credentials from the save file.
 *
 * @return {Promise<OAuth2Client|null>}
 */
async function loadSavedCredentialsIfExist() {
    try {
      const content = await fs.readFile(TOKEN_PATH);
      const credentials = JSON.parse(content);

      // 如果文件中包含 refresh_token，直接使用它創建 OAuth2Client
      if (credentials.refresh_token) {
        const oauth2Client = new google.auth.OAuth2(
          credentials.client_id,
          credentials.client_secret,
          // YOUR_REDIRECT_URI
        );
        oauth2Client.setCredentials({
          refresh_token: credentials.refresh_token
        });
        return oauth2Client;
      }
      
      // 如果沒有 refresh_token，則使用之前的方法
      return google.auth.fromJSON(credentials);
    } catch (err) {
      console.log('[loadSavedCredentialsIfExist]' + err);
      logger.error('[loadSavedCredentialsIfExist Error] ' + err);
      return null;
    }
  }
  
  /**
   * Serializes credentials to a file compatible with GoogleAuth.fromJSON.
   *
   * @param {OAuth2Client} client
   * @return {Promise<void>}
   */
  async function saveCredentials(client) 
  {
    try {
    const content = await fs.readFile(CREDENTIALS_PATH);
    const keys = JSON.parse(content);
    const key = keys.installed || keys.web;
    const payload = JSON.stringify({
      type: 'authorized_user',
      client_id: key.client_id,
      client_secret: key.client_secret,
      refresh_token: client.credentials.refresh_token,
    });
    await fs.writeFile(TOKEN_PATH, payload);
  } catch (err) {
    console.error('[saveCredentials]' + err);
    logger.error('[saveCredentials Error] ' + err);
    return null;
  }
  }
  
  /**
   * Load or request or authorization to call APIs.
   *
   */
  async function authorize() {
    logger.debug(`[authorize]start`);   
    let client = await loadSavedCredentialsIfExist();
    if (client) { 
      try {
        // 嘗試刷新 access token
        await client.getAccessToken();
        // 如果成功，保存新的憑證
        await saveCredentials(client);
      } catch (err) {
        console.error('Error refreshing access token:', err);
        logger.error('[authorize Error] Error refreshing access token: ' + err);
        // 如果刷新失敗，將 client 設為 null，以便重新授權
        client = null;
      }
    }
    
    if (!client) {
      // 如果無法使用 refresh token，則需要重新進行完整的授權流程
      client = await authenticate({
        scopes: SCOPES,
        keyfilePath: CREDENTIALS_PATH,
      });
      if (client.credentials) {
        await saveCredentials(client);
      }
    }

    logger.debug(`[authorize]end`);   
    return client;
  }
  
  /**
   * Lists the next 10 events on the user's primary calendar.
   * @param {google.auth.OAuth2} auth An authorized OAuth2 client.
   */
  async function listEvents(auth) {
    const calendar = google.calendar({version: 'v3', auth});
    const res = await calendar.events.list({
      calendarId: 'primary',
      timeMin: new Date().toISOString(),
      maxResults: 10,
      singleEvents: true,
      orderBy: 'startTime',
    });
    const events = res.data.items;
    if (!events || events.length === 0) {
      console.log('No upcoming events found.');
      return;
    }
    console.log('Upcoming 10 events:');
    events.map((event, i) => {
      const start = event.start.dateTime || event.start.date;
      console.log(`${start} - ${event.summary}`);
    });
  }

  async function checkSameEventIsExist(calendarId, query, startDate, endDate) {
    try{      
      let auth = await authorize();
      const calendar = google.calendar({version: 'v3', auth});
      const resp = await calendar.events.list({
        calendarId: calendarId,
        q: query,
        timeMin: startDate.toISOString(),
        timeMax: endDate.toISOString(),
        singleEvents: true,
        orderBy: 'startTime',
      });
      return resp.data.items.length > 0;
    }catch(err){
      console.error('[checkSameEventIsExist]' + err);
      logger.error('[checkSameEventIsExist Error] ' + err);
    }
  }

  /**
   * 行事曆新增事件
   *
   * @param {CalendarEvent} event 
   */
  async function InsertEvents(event) {
    try{
    let auth = await authorize();
    const calendar = google.calendar({version: 'v3', auth});
      let response = await calendar.events.insert({
        auth: auth,
        calendarId: 'primary',
        resource: event,
      })  
      let data = response.data;
      console.log('Event created: %s', data.htmlLink);
      logger.info(`[InsertEvents Success]行事曆事件【${data.summary}】已建立!${data.htmlLink}`);
    }catch(err){
      console.error('There was an error contacting the Calendar service: ' + err);
      logger.error('[InsertEvents Error]There was an error contacting the Calendar service: ' + err);  
    }
  }

  async function TestInsertEvent(dateTime){    
  InsertEvents(new CalendarEvent({
      summary: 'Hello World; first Calendar，去丟垃圾',
      location: '台北市士林區中山北路 122號',
      description: '我是一段描述，去丟垃圾',
      start: {
        dateTime: dateTime,
        timeZone: 'Asia/Taipei'
      },
      end: {
        dateTime: dateTime,
        timeZone: 'Asia/Taipei'
      },
      recurrence: [
      ],
      attendees: [],
      reminders: {
        useDefault: false,
        overrides: [
          { method: 'email', minutes: 24 * 60 },
          { method: 'popup', minutes: 30 }
        ]
      }
    }));
  }  

//#endregion

