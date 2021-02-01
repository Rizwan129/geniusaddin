/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.initialize=function() {
  
    // document.getElementById("logout_user").onclick = logout;
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    var code=getParameterByName("access_token");
    console.log(code);
    if(code)
    {
      Office.context.ui.messageParent(code);

    }
    else
    {

      Office.context.ui.messageParent("Sign-Out");

    }
   
  
};

//Get Token from Query String
function getParameterByName(name, url = window.location.href) {
  name = name.replace(/[\[\]]/g, '\\$&');
  var regex = new RegExp('[?&#]' + name + '(=([^&#]*)|&|#|$)'),
      results = regex.exec(url);
  if (!results) return null;
  if (!results[2]) return '';
  return decodeURIComponent(results[2].replace(/\+/g, ' '));
}

// var loginDialog;
//  function logout(){


//   Office.context.ui.displayDialogAsync("https://g.auth.us-east-1.amazoncognito.com/logout?client_id=635u3pcfjo3bl7fb0p1g5269bl&response_type=token&redirect_uri=https://localhost:3000/&identity_provider=Intuit&scope=aws.cognito.signin.user.admin%20openid",
//   { height: 70, width: 25 }, function (result) {
//       if (result.error)
//           console.log(result.error.message + ' ' + result.error.code);
//       loginDialog = result.value;
//       loginDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived,processMessage);
//   });  
// }

// var processMessage = function (arg) {
//   loginDialog.close();
// }