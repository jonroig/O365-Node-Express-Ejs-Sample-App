/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */

// routes/mail.js

var url = require('url');
var request = require('request');
var appSettings = require('../models/appSettings.js');


module.exports = function (app, passport, utils) {
    app.use('/mail', function (req, res, next) {
        passport.getAccessToken(appSettings.resources.exchange, req, res, next);
    })

    // Get a messaget list in the user's Inbox using the O365 API,
    // displaying To, Subject and Preview for each message.
    app.get('/mail', function (req, res, next) {
        request.get(
            appSettings.apiEndpoints.exchangeBaseUrl + "/messages",
            { auth : { 'bearer' : "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Ik1uQ19WWmNBVGZNNXBPWWlKSE1iYTlnb0VLWSIsImtpZCI6Ik1uQ19WWmNBVGZNNXBPWWlKSE1iYTlnb0VLWSJ9.eyJhdWQiOiJodHRwczovL2FwaS5vZmZpY2UuY29tL2Rpc2NvdmVyeS8iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC83ZTYxNzhjZi02ZTg0LTQyY2EtOTIwNi01Nzc5NTg1ZWMyMzcvIiwiaWF0IjoxNDM3NjE3NzY1LCJuYmYiOjE0Mzc2MTc3NjUsImV4cCI6MTQzNzYyMTY2NSwidmVyIjoiMS4wIiwidGlkIjoiN2U2MTc4Y2YtNmU4NC00MmNhLTkyMDYtNTc3OTU4NWVjMjM3IiwiZW1haWwiOiJqcm9pZ0Bnb2RhZGR5LmNvbSIsImlkcCI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0L2Q1ZjE2MjJiLTE0YTMtNDVhNi1iMDY5LTAwM2Y4ZGM0ODUxZi8iLCJhbHRzZWNpZCI6IjU6OjEwMDNCRkZEODU0MTQ1QzkiLCJzdWIiOiJlMUF4TGFWTkwzNHlCa1BFZWdzc3c1WDZWSnN2UHFwSVVkS3p2b2RJSmo4IiwiZ2l2ZW5fbmFtZSI6IkpvbmF0aGFuIiwiZmFtaWx5X25hbWUiOiJSb2lnIiwibmFtZSI6IkdvIERhZGR5LmNvbSwgTExDLCBBIERlbGF3YXJlIENvbXBhbnkiLCJhbXIiOlsicHdkIl0sInVuaXF1ZV9uYW1lIjoianJvaWdAZ29kYWRkeS5jb20iLCJhcHBpZCI6IjdkMmVlYjA5LWZmZGYtNDg2Mi04NGMzLTBjMGRiMGQ4YTNhYSIsImFwcGlkYWNyIjoiMSIsInNjcCI6IkNhbGVuZGFycy5SZWFkV3JpdGUgRGlyZWN0b3J5LkFjY2Vzc0FzVXNlci5BbGwgRGlyZWN0b3J5LlJlYWRXcml0ZS5BbGwgRmlsZXMuUmVhZCBNYWlsLlJlYWRXcml0ZSBNYWlsLlNlbmQgb2ZmbGluZV9hY2Nlc3Mgb3BlbmlkIFRhc2tzLlJlYWQuQWxsIFVzZXIuUmVhZCBVc2VyLlJlYWQuQWxsIFVzZXIuUmVhZFdyaXRlIFVzZXIuUmVhZFdyaXRlLkFsbCBVc2VyUHJvZmlsZS5SZWFkIiwiYWNyIjoiMSJ9.AVAkDtBlQYiHCxfRDEc6WufxhAxzQ2nqwkv2FqFlfCVVN2liW2PPDkJkt4sSFrKbEKZ2nKZmrsxj6aRuqa5jZU0e0vK7j_IEOVGnMdQa57yZ4zoLRNSDKRB3mjV7KNORJ5lgQ062I50uYDQHMFI0yx5SYHeTOhz8ithOvgf8OsTN3F9CX3ACSWpms9giVSnDV_xMlyH9eLSPbtHxVJXYoNdOzxVhKc3SqHTX0tDuqgNB1rlHQ9PH--CwQl2_El5zqAOiBW5iLYLdmV-sTUAi24XZIaKwNOH9FmSxTb5KBnXhyhyBHcjXh333AN3HqsleLUqAe83SQMZArmPHoaBuiQ" } },
            function (error, response, body) {
                if (error) {
                    next(error);
                }
                else {
                    data = { user: passport.user, msgs: JSON.parse(body)['value'] };
                    res.render('mail', { data: data });
                }
            }
        );
    });

    // GET a given message and display content to the user,
    // displaying the message as-is in HTML.
    app.get('/mail/view', function (req, res, next) {
        var id = url.parse(req.url, true).query.id;
        request.get(
            appSettings.apiEndpoints.exchangeBaseUrl + "/messages/" + id,
            { auth : { 'bearer' : passport.user.getToken(appSettings.resources.exchange).access_token } },
            function (error, response, body) {
                if (error) {
                    next(error);
                }
                else {
                    var jsonBody = JSON.parse(body);
                    res.end(jsonBody.Body.Content);
                }
            }
        );
    })

    // delete a selected email message using the O365 API.
    app.get('/mail/delete', function (req, res, next) {
        var id = url.parse(req.url, true).query.id;
        request.del(
            appSettings.apiEndpoints.exchangeBaseUrl + "/messages/" + id,
            { auth : { 'bearer' : passport.user.accessToken } },
            function (error, response, body) {
                if (error) {
                    next(error);
                }
                else {
                    res.redirect('/mail')
                }
            }
        );

    })

    // Pop up a message editor for th user to add comment to the original message.
    app.get('/mail/reply', function (req, res, next) {
        var id = url.parse(req.url, true).query.id;
        request.get(
            appSettings.apiEndpoints.exchangeBaseUrl + "/messages/" + id,
            { auth : { 'bearer' : passport.user.getToken(appSettings.resources.exchange).access_token } },
            function (error, response, body) {
                if (error) {
                    next(error);
                }
                else {
                    var jsonBody = JSON.parse(body);
                    var content = utils.htmlText(jsonBody.Body.Content);
                    res.render('mailReply', {
                        user : passport.user,
                        messageId : id,
                        recipients: jsonBody.Sender.EmailAddress.Address,
                        subject: jsonBody.Subject,
                        content: content
                    });
                }
            }
        );
    });

    // Pop up a message composer for the user to creat and send a new mail message.
    app.get('/mail/new', function (req, res, next) {
        res.render('mailedit', { user: passport.user, recipients : "user@domain", subject: "Test", content : "" });
    })

    // send a new mail message to a specific recipient using O365 API.
    // The request body must be a JSON string, not an JSON object.
    app.post('/mail/send', function (req, res, next) {
        var reqBody = {
            'Message' : {
                'Subject': req.body.subject,
                'Body': { 'ContentType': "Text", 'Content': req.body.message },
                'ToRecipients' : [{ 'EmailAddress': { 'Address' : req.body.to } }]
            },
            'SaveToSentItems' : 'false'
        };
        var reqHeaders = { 'content-type': 'application/json'};
        var reqUrl = appSettings.apiEndpoints.exchangeBaseUrl + "/sendmail";
        var reqAuth = { 'bearer': passport.user.getToken(appSettings.resources.exchange).access_token };

        request.post(
            { url: reqUrl, headers: reqHeaders, body: JSON.stringify(reqBody), auth: reqAuth },
            function (err, response, body) {
                if (err) { next(err); }
                else {
                    if (response.statusCode == 403) {
                        err = { status : response.statusCode, msg : body , stack : "Failed to send mail to " + req.body.to };
                        next(err);
                    }
                    else {
                        res.redirect('/mail' );
                    }

                }
            }
        );
    })

    // reply a mail message using the O365 API. The app-submitted request body
    // contains only the reply.The API will include the original message in the
    // final request body before sending it over the wire.
    app.post('/mail/reply', function (req, res, next) {
        var messageId = req.body.messageId;

        var reqBody = { 'Comment' : req.body.comment };
        var reqHeaders = { 'content-type': 'application/json' };
        var reqUrl = appSettings.apiEndpoints.exchangeBaseUrl + "/messages/" + messageId + "/reply";
        var reqAuth = { 'bearer': passport.user.getToken(appSettings.resources.exchange).access_token };

        request.post(
            { url: reqUrl, headers: reqHeaders, body: JSON.stringify(reqBody), auth: reqAuth },
            function (err, response, body) {
                if (err) { next(err); }
                else {
                    if (response.statusCode == 403) {
                        err = { status : response.statusCode, msg : body , stack : "Failed to reply mail to " + req.body.to };
                        next(err);
                    }
                    else {
                        res.redirect('/mail');
                    }
                }
            }
        );
    })

}

// *********************************************************
//
// O365-Node-Express-Ejs-Sample-App, https://github.com/OfficeDev/O365-Node-Express-Ejs-Sample-App
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************

