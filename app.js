/*
 * Copyright (c) Microsoft All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

var auth = require('./auth');
var graph = require('./graph');

// Get an access token for the app.
auth.getAccessToken().then(function (token) {
  // Get all of the users in the tenant.

  // Create an event on each user's calendar.
  graph.getListItems(token)
    .then(function (items) {
      items.forEach(function (item) {
        graph.getListItem(token, item.id).then(function (itemInfo) {
          console.log('%c %c', itemInfo.Title, itemInfo.VTXAnswer);
        }, function (error) {
          console.error('>>> Error getting items: ' + error);
        });
      });
    }, function (error) {
      console.error('>>> Error getting items: ' + error);
    });
}, function (error) {
  console.error('>>> Error getting access token: ' + error);
});
