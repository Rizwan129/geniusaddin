import moment from './moment.js';
import {getValue} from '../utils.js';

Date.prototype.yyyymmdd = function() {
  var mm = this.getMonth() + 1; // getMonth() is zero-based
  var dd = this.getDate();

  return [this.getFullYear(),
          (mm>9 ? '' : '0') + mm,
          (dd>9 ? '' : '0') + dd
         ].join('-');
};

function ExcelDateToJSDate(serial) {
  var utc_days  = Math.floor(serial - 25569);
  var utc_value = utc_days * 86400;                                        
  var date_info = new Date(utc_value * 1000);

  var fractional_day = serial - Math.floor(serial) + 0.0000001;

  var total_seconds = Math.floor(86400 * fractional_day);

  var seconds = total_seconds % 60;

  total_seconds -= seconds;

  var hours = Math.floor(total_seconds / (60 * 60));
  var minutes = Math.floor(total_seconds / 60) % 60;

  return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);
}

// return date in a string format
function parseDateIfObject(date){
  if (Object.keys(date).length < 7){
    date = ExcelDateToJSDate(Object.values(date).join('')); // JS object format
    date = date.yyyymmdd();
  } // else pass as it is
  return date;
}

//TODO change url to valid end point
var GENIUSSHEETS_URL = 'https://geniussheets.herokuapp.com';
  /**
   * Calculates the total expense under given category from startDate till endDate inclusive.
   * @customfunction
   * @param {string}  category The category of Expense. Keep empty to fetch total expense. Click on List Categories in GeniusSheets menu to see possible categories.
   * @param {string} startDate The valid starting date in format "yyyy-mm-dd" or in format "MMM-yyyy" or a cell address containg a date.
   * @param {string} [endDate] endDate The valid ending date in format "yyyy-mm-dd" or in format "MMM-yyyy" or a cell address containg a date.
   * @return {number} The total expense
   */
  async function EXPENSE(category, startDate, endDate) {
    startDate = parseDateIfObject(startDate);
    endDate = parseDateIfObject(endDate);

    if(category === ""){
      category = "TOTAL";
    }
    if(endDate === null){
      endDate = startDate;
    }
    console.log(category+' '+startDate+' '+endDate);
    try {
        var [startDate_,endDate_] = getValidDate(startDate,endDate);
    }catch (e) {
        console.log(e)
        throw e;
    }
    let access_token = await getValue('access_token');
    if(access_token!==null){
        let url = GENIUSSHEETS_URL + '/api/getCategoryAmount/Expense/' + category + '/'
          + startDate_ + '/' + endDate_ + '/';
      return getData(url,access_token);

    } else {
      console.log('Not Logged In');
      throw new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable,'User not logged in');
    }

  }

/**
 * Calcutates the total Revenue under given cateory from startDate till endDate inclusive. Click on Function Details in GeniusSheets menu for more information.
 * @customfunction
 * @param {string} category The category of Revenue. Keep empty to fetch total revenue. Click on List Categories in GeniusSheets menu to see possible categories.
 * @param {string} startDate The valid starting date in format "yyyy-mm-dd" or in format "MMM-yyyy" or a cell address containg a date.
 * @param {string} [endDate] endDate The valid ending date in format "yyyy-mm-dd" or in format "MMM-yyyy" or a cell address containg a date.
 * @return {number} The total Revenue
 */

async function REVENUE(category, startDate, endDate) {
  startDate = parseDateIfObject(startDate);
  endDate = parseDateIfObject(endDate);

  if(category === ""){
    category = "TOTAL";
  }
  if(endDate === null){
    endDate = startDate;
  }
  console.log(category+' '+startDate+' '+endDate);
    try {
        var [startDate_,endDate_] = getValidDate(startDate,endDate);
    }catch (e) {
        console.log(e);
        throw e;
    }
  let access_token = await getValue('access_token');
  if(access_token!==null){
    let url = GENIUSSHEETS_URL + '/api/getCategoryAmount/Income/' + category + '/'
        + startDate_ + '/' + endDate_ + '/';
    return getData(url,access_token);

  } else {
    console.log('Not Logged In');
    throw new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable,'User not logged in');
  }

}

/**
 * Calcutates the total Cost of Goods Sold under given cateory from startDate till endDate inclusive. Click Function Details in GeniusSheets menu for more information.
 * @customfunction
 * @param {string} category The category of Cost of Goods Sold. Keep empty to fetch total amount. Click on List Categories in GeniusSheets menu to see possible categories.
 * @param {string} startDate The valid starting date in format "yyyy-mm-dd" or in format "MMM-yyyy" or a cell address containg a date.
 * @param {string} [endDate] endDate The valid ending date in format "yyyy-mm-dd" or in format "MMM-yyyy" or a cell address containg a date.
 * @return {number} The total expense
 */
async function COGS(category, startDate, endDate) {
  startDate = parseDateIfObject(startDate);
  endDate = parseDateIfObject(endDate);

    if(category === ""){
        category = "TOTAL";
    }
    if(endDate === null){
        endDate = startDate;
    }
    console.log(category+' '+startDate+' '+endDate);
    try {
        var [startDate_,endDate_] = getValidDate(startDate,endDate);
    }catch (e) {
        console.log(e)
        throw e;
    }
    let access_token = await getValue('access_token');
    if(access_token!==null){
        let url = GENIUSSHEETS_URL + '/api/getCategoryAmount/COGS/' + category + '/'
            + startDate_ + '/' + endDate_ + '/';
        return getData(url,access_token);

    } else {
        console.log('Not Logged In');
        throw new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable,'User not logged in');
    }

}

/**
 * Calcutates the total Other Expense in under given cateory from startDate till endDate inclusive. Click Function Details in GeniusSheets menu for more information.
 * @customfunction
 * @param {string} category The category of Other Expense. Keep empty to fetch total amount. Click on List Categories in GeniusSheets menu to see possible categories.
 * @param {string} startDate The valid starting date in format "yyyy-mm-dd" or in format "MMM-yyyy" or a cell address containg a date.
 * @param {string} [endDate] endDate The valid ending date in format "yyyy-mm-dd" or in format "MMM-yyyy" or a cell address containg a date.
 * @return {number} The total Revenue
 */

async function OTHEREXPENSE(category, startDate, endDate) {
  startDate = parseDateIfObject(startDate);
  endDate = parseDateIfObject(endDate);

    if(category === ""){
        category = "TOTAL";
    }
    if(endDate === null){
        endDate = startDate;
    }
    console.log(category+' '+startDate+' '+endDate);
    try {
        var [startDate_,endDate_] = getValidDate(startDate,endDate);
    }catch (e) {
        console.log(e);
        throw e;
    }
    let access_token = await getValue('access_token');
    if(access_token!==null){
        let url = GENIUSSHEETS_URL + '/api/getCategoryAmount/OtherExpenses/' + category + '/'
            + startDate_ + '/' + endDate_ + '/';
        return getData(url,access_token);

    } else {
        console.log('Not Logged In');
        throw new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable,'User not logged in');
    }

}

/**
 * Calcutates the total Other Income under given cateory from startDate till endDate inclusive. Click Function Details in GeniusSheets menu for more information.
 * @customfunction
 * @param {string} category The category of Other Income. Keep empty to fetch total amount. Click on List Categories in GeniusSheets menu to see possible categories.
 * @param {string} startDate The valid starting date in format "yyyy-mm-dd" or in format "MMM-yyyy" or a cell address containg a date.
 * @param {string} [endDate] endDate The valid ending date in format "yyyy-mm-dd" or in format "MMM-yyyy" or a cell address containg a date.
 * @return {number} The total Revenue
 */

async function OTHERINCOME(category, startDate, endDate) {
  startDate = parseDateIfObject(startDate);
  endDate = parseDateIfObject(endDate);

    if(category === ""){
        category = "TOTAL";
    }
    if(endDate === null){
        endDate = startDate;
    }
    console.log(category+' '+startDate+' '+endDate);
    try {
        var [startDate_,endDate_] = getValidDate(startDate,endDate);
    }catch (e) {
        console.log(e);
        throw e;
    }
    let access_token = await getValue('access_token');
    if(access_token!==null){
        let url = GENIUSSHEETS_URL + '/api/getCategoryAmount/OtherIncome/' + category + '/'
            + startDate_ + '/' + endDate_ + '/';
        return getData(url,access_token);

    } else {
        console.log('Not Logged In');
        throw new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable,'User not logged in');
    }

}

function getValidDate(startDate,endDate){
    var startDate_ = convertToDateString_(startDate,true);
    var endDate_ = convertToDateString_(endDate,false);

    if (!validateDate_(startDate_,true)) {

        throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue,
            'Parameter startDate expects a valid date string in form "yyyy-mm-dd" or a date cell address. Also startDate should not be in future');
    }
    if (!validateDate_(endDate_,false)) {
        throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue,
            'Parameter endDate expects a valid date string in form "yyyy-mm-dd" or a date cell address.');
    }
    if (startDate_ > endDate_) {
        throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue,
            'startDate should not be after endDate');
    }
    return [startDate_,endDate_]
}

async function getData(url,access_token){
  var headers = {
    'Authorization' : 'Token ' +access_token
  };
  var options = {
    'method' : 'GET',
    'contentType': 'application/json',
    'headers': headers,
    'muteHttpExceptions' : true
  };

  const response = await fetch(url,options);
  if(response.ok ){
    const jsonResponse = await response.json();
    return jsonResponse.totalAmount;
  }
  else if (response.status === 401) {
    throw new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable,'Quickbooks authentication unavailable. Initiate Quickbooks Sign in inside GeniusSheets');
  }
  else if (response.status === 400) {
    throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue,'Invalid startDate or endDate. Make sure the dates are valid.');
  }
  else if (response.getResponseCode() !== 200) {
    throw new Error('Internal Server Error. Contact GeniusSheets');
  }
}

function getNumberforMonth(Month) {
    const months = {
        jan: "01",
        feb: "02",
        mar: "03",
        apr: "04",
        may: "05",
        jun: "06",
        jul: "07",
        aug: "08",
        sep: "09",
        oct: "10",
        nov: "11",
        dec: "12"
    }
    return months[Month] != null;
}

function isNumeric_(c) {
    return c >= '0' && c <= '9';
}


/**
 * Function to convert date object in form "yyyy-mm-dd"
 * @param {string} date Date
 * @param BoolStart
 */
function convertToDateString_(date, BoolStart) {
    console.log(date, date.length, typeof date);
    if (typeof date === "string" && date.length > 0 && !isNumeric_(date[0])) {
        //check if string month exists in our list
        if (getNumberforMonth(date.toLowerCase().split("-")[0]) === false) {
            return date;
        }

        if (date.length === 6) {
            console.log("Received small length date : ", date);
            date = date.slice(0, 4) + '20' + date.slice(4, 6);
            console.log("Converted to :", date);
        }

        console.log("trying to convert to YYYY-MM-DD string :", date);
        try {
            if (BoolStart) {
                return moment(date).startOf('month').format("YYYY-MM-DD")
            } else {
                return moment(date).endOf('month').format("YYYY-MM-DD")
            }
        } catch (e) {
            console.log(e);
            return date;
        }
    }
    if (typeof date === "object") {
        return moment(date).format("YYYY-MM-DD");
    }
    return date;
}


/**
 * Function to check if a string in form "yyyy-mm-dd" and valid
 * @param {string} date Date
 * @param {boolean} isStartDate to validate start and end date
 */
function validateDate_(date,isStartDate) {
    if (typeof date == "string") {
        var curDate = convertToDateString_(new Date());
        var dateRegex = /^\d{4}[-]\d{2}[-]\d{2}$/
        var res = Date.parse(date);
        if (isStartDate)
            return (!isNaN(res) && date <= curDate && date.match(dateRegex));
        else
            return (!isNaN(res) && date.match(dateRegex));
    }
    return false;
}
