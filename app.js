/*
 * Copyright (c) Microsoft All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

var auth = require('./auth');
var graph = require('./graph');
console.log('hi');

// Get an access token for the app.
auth.getAccessToken().then(function (token) {
  // Get all of the users in the tenant.
  graph.getUsers(token)
    .then(function (/* users*/) {
      // Create an event on each user's calendar.
      graph.getEvents(token, 'chandima@chandimaranaweera.com')
        .then(function (events) {
          events.forEach(function (event) {
            console.log('%c %c', event.start.dateTime, event.end.dateTime);
          });
        });
    }, function (error) {
      console.error('>>> Error getting users: ' + error);
    });
}, function (error) {
  console.error('>>> Error getting access token: ' + error);
});
