/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

// This sample uses an open source OAuth 2.0 library that is compatible with the Azure AD v2.0 endpoint. 
// Microsoft does not provide fixes or direct support for this library. 
// Refer to the library’s repository to file issues or for other support. 
// For more information about auth libraries see: https://azure.microsoft.com/documentation/articles/active-directory-v2-libraries/ 
// Library repo: https://github.com/MrSwitch/hello.js

"use strict";

(function () {
  angular
    .module('app')
    .service('GraphHelper', ['$http', function ($http) {
      console.log("here");
      console.log(graphScopes);
      console.log(clientId);

      // Initialize the auth request.
      hello.init( {
        aad: clientId // from public/scripts/config.js
        }, {
        redirect_uri: redirectUrl,
        scope: graphScopes
      });

      return {

        // Sign in and sign out the user.
        login: function login() {
          hello('aad').login({
            display: 'page',
            state: 'abcd'
          });
        },
         logout: function logout() {
          hello('aad').logout();
          delete localStorage.auth;
          delete localStorage.user;
        },

        // Get the profile of the current user.
        me: function me() {
          return $http.get('https://graph.microsoft.com/v1.0/me');
        },

        // Send an email on behalf of the current user.
        sendMail: function sendMail(email) {
          return $http.post('https://graph.microsoft.com/v1.0/me/sendMail', { 'message' : email, 'saveToSentItems': true });        
        },
        getFile: function getFile(file) {
          var path = "test";
          var driveId = "b!YW8EejK3B0uwtyl1xobCU48Z1nPPMMBMslrknikGiPL0ShoWhRZYTIJzaH8j3x6u";
          var driveItem = "01EUKMIKZ2F6RTVZICTNDKU7DQ5ONJIX6N";
          var remoteItem = "01EUKMIKZ2F6RTVZICTNDKU7DQ5ONJIX6N";
          var sharedDriveId = "b!jGlmx2fc-EW8VVoTSHnTKy7q06Bl5WBMh7aziPkw1t8MatZpvwlvSZJzVL-0-R-y";
          var myDriveId = "b!YW8EejK3B0uwtyl1xobCU48Z1nPPMMBMslrknikGiPL0ShoWhRZYTIJzaH8j3x6u";
          var uniqueId = "{3AA32F3A-02E5-469B-AA7C-70EB9A945FCD}";
          var listId = "69d66a0c-09bf-496f-9273-54bfb4f91fb2";
          var siteId = "c766698c-dc67-45f8-bc55-5a134879d32b";
          var webId = "a0d3ea2e-e565-4c60-87b6-b388f930d6df";
          var excelItemId = "01EUKMIK5KO6M7I2IWTBBYYMTNW2LB5CTP";
          var worksheetName = "Survey1";

//          return $http.get('https://graph.microsoft.com/v1.0/me/drive/root/children');
          return $http.get('https://graph.microsoft.com/v1.0/drives/' + sharedDriveId + '/items/' + excelItemId + '/workbook/worksheets/' + worksheetName + '/tables/Table1/rows' );

        }
      }
    }]);
})();