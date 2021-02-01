import {saveProp,getCategories} from "../utils";
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.initialize=function(){
  
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("run").onclick = run;
    document.getElementById("loginform").onclick = Loginfn; //login function
    document.getElementById("logout").onclick = logout;
    document.getElementById("refreshData").onclick = refreshData;
    document.getElementById("loginformaws").onclick = awsLogin;
    
    // displays form if no access_token present in cache
   
      document.getElementById("app-body").style.display = "flex";
     
      document.getElementById("message").style.display = "block";   
   
    //Get categories function to show available categoies can be used in 
    // console.log(getCategories())
  
}


export async function getValue(key) {
  let value = await OfficeRuntime.storage.getItem(key);
  return value;
}


export async function logout(){

    OfficeRuntime.storage.removeItem("access_token")
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("message").style.display = "none";



  if(getValue("access_token") != null)
  {
    OfficeRuntime.storage.removeItem("access_token")
  }
  // else{
  //   // run()
  // }
}

export async function refreshData(){
  try {
    await Excel.run(async context => {
      context.workbook.application.calculate('Full');
      return context.sync(); 
      // var sheet = context.workbook.worksheets.getActiveWorksheet();
      // var foundRanges = sheet.findAll("=GENIUS.", {
      //     completeMatch: false, // findAll will match the whole cell value
      //     matchCase: false // findAll will not match case
      // });

      // return context.sync()
      //     .then(function() {
      //         var tmp = foundRanges.formulas;
      //         foundRanges.formulas = null;
      //         foundRanges.formulas = tmp;
      //         // foundRanges.format.fill.color = "green"
      // });
    });
  } catch (error) {
    console.error(error);
  }
}
//login user
var loginDialog;
export async function awsLogin(){
  Office.context.ui.displayDialogAsync("https://g.auth.us-east-1.amazoncognito.com/oauth2/authorize?client_id=635u3pcfjo3bl7fb0p1g5269bl&response_type=token&redirect_uri=https://localhost:3000/taskpane.html&identity_provider=Intuit&scope=aws.cognito.signin.user.admin%20openid",
  { height: 70, width: 25 }, function (result) {
      if (result.error)
          console.log(result.error.message + ' ' + result.error.code);
      loginDialog = result.value;
      loginDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived,processMessage);
  });  
}

 var processMessage = function (arg) {
  loginDialog.close();
  var messageResult = arg.message;
var resultJWT=  parseJwt (messageResult);
localStorage.setItem("userid",resultJWT.username);
window.location.href="/userInfo.html";
}

//get Information from Token
function parseJwt (token) {
  var base64Url = token.split('.')[1];
  var base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
  var jsonPayload = decodeURIComponent(atob(base64).split('').map(function(c) {
      return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
  }).join(''));

  return JSON.parse(jsonPayload);
}
// @description :
// gets input values
// validates input values
// hits API
// set object in session
// check if session then displays Logout
// else shows the form
export async function Loginfn() {
  // let obj={}
  let username="";
  let password ="";
  var GENIUSSHEETS_URL = 'https://geniussheets.herokuapp.com';

  username = document.getElementById("username").value;
  password =  document.getElementById("inputpassword").value;
 
  document.getElementById("error").style.display = "none";
  document.getElementById("signinerror").style.display = "none";

  document.getElementById("loginform").style.display = "none";
  document.getElementById("loginformload").style.display = "block";



  if(username.length == 0 || password.length == 0)
  {
    // run();
    document.getElementById("loginform").style.display = "block";
    document.getElementById("loginformload").style.display = "none";
    document.getElementById("error").style.display = "flex";
    return;  
  }
  // run()

  // HIT API
  var data = {
    username : username,
    password : password
  }
  var myHeaders = new Headers();
  myHeaders.append("Content-Type", "application/json");
  var options = {
    'method' : 'post',
    'body' : JSON.stringify(data),
    'muteHttpExceptions' : true,
    redirect: 'follow',
    headers: myHeaders,
  };
  const response = await fetch(GENIUSSHEETS_URL+'/api/login/',options);
  console.log(response);
  if(response.ok)
  {
    // var response_obj = JSON.parse(response.getContentText());    
    const response_obj = await response.json();
    OfficeRuntime.storage.setItem("access_token", response_obj.access_token);
    document.getElementById("loginform").style.display = "block";
    document.getElementById("loginformload").style.display = "none";
    document.getElementById("app-body").style.display = "none";
    document.getElementById("message").style.display = "block";        
  }else{
    document.getElementById("loginform").style.display = "block";
    document.getElementById("loginformload").style.display = "none";
    document.getElementById("signinerror").style.display = "block";
  }

  // test code
  // document.getElementById("app-body").style.display = "none";
  // document.getElementById("message").style.display = "block"; 
  // run()
}


// unused test function for refrences
export async function run() {
  try {
    await Excel.run(async context => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

