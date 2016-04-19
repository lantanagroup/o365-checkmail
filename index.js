#! /usr/bin/env node

// Initialize the OAuth2 Library
var request = require('request-promise');
var _ = require('underscore');
var Q = require('q');
var xmlbuilder = require('xmlbuilder');
var argv = require('yargs')
    .option('username', {
        alias: 'u',
        require: true
    })
    .option('password', {
        alias: 'p',
        require: true
    })
    .option('site', {
        alias: 's',
        require: true,
        describe: 'The url to the client application site (ex: https://login.microsoftonline.com/TENANT_ID)'
    })
    .option('id', {
        alias: 'i',
        require: true,
        describe: 'The id of the client application'
    })
    .option('secret', {
        alias: 'x',
        require: true,
        describe: 'The secret (key) for the client application'
    })
    .option('tokenPath', {
        alias: 't',
        default: '/oauth2/token'
    })
    .option('authorizationPath', {
        alias: 'a',
        default: '/oauth2/auth'
    })
    .option('resource', {
        alias: 'r',
        default: 'https://graph.microsoft.com'
    })
    .argv;

var credentials = {
    clientID: argv.id,
    clientSecret: argv.secret,
    site: argv.site,
    tokenPath: argv.tokenPath,
    authorizationPath: argv.authorizationPath,
    tokenConfig: {
        username: argv.username,
        password: argv.password,
        resource: argv.resource
    }
};

var token;

var oauth2 = require('simple-oauth2')(credentials);
oauth2.password.getToken(credentials.tokenConfig, function saveToken(error, result) {
    if (error) {
        console.log('Access Token Error', error.message);
        return process.exit(1);
    }
    token = oauth2.accessToken.create(result);

    // https://graph.microsoft.io/en-us/docs
    var foldersRequestOptions = {
        method: 'GET',
        url: 'https://graph.microsoft.com/v1.0/me/mailFolders',
        headers: {
            'Authorization': 'Bearer ' + token.token.access_token,
            'Accept': 'application/json'
        },
        json: true
    };

    request(foldersRequestOptions)
        .then(function(foldersResponse) {
            var inboxFolder = _.find(foldersResponse.value, function(folder) {
                return folder.displayName == 'Inbox';
            });

            var msgRequestOptions = {
                method: 'GET',
                url: 'https://graph.microsoft.com/v1.0/me/MailFolders/inbox/messages?$top=300&$filter=(isRead eq false)&$select=createdDateTime,subject',
                headers: {
                    'Authorization': 'Bearer ' + token.token.access_token,
                    'Accept': 'application/json'
                },
                json: true
            };

            return request(msgRequestOptions);
        })
        .then(function(msgResponse) {
            var output = xmlbuilder.create('root');
            output.ele('count', msgResponse.value.length);

            _.each(msgResponse.value, function(msg) {
                var messageEle = output.ele('message');
                messageEle.ele('subject', msg.subject);
                messageEle.ele('createdDateTime', msg.createdDateTime);
                messageEle.ele('id', msg.id);
            });

            var outputXml = output.end({ pretty: true });
            console.log(outputXml);
            process.exit(0);
        })
        .catch(function(err) {
            console.log('Error retrieving data from Office365: ' + err);
            return process.exit(1);
        });
});