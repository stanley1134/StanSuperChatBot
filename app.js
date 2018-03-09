/*-----------------------------------------------------------------------------
A simple echo bot for the Microsoft Bot Framework. 
-----------------------------------------------------------------------------*/

var dt = require('./data.js');
var $ = require('jQuery');
var getJSON = require('get-json');
var cardObj = require('./card.js');
var weathercardObj = require('./weathercard.js');

var restify = require('restify');
var builder = require('botbuilder');
var botbuilder_azure = require("botbuilder-azure");

var Connection = require('tedious').Connection;
var Request = require('tedious').Request;
var csomapi = require('csom-node');
var requestpromise = require('request-promise').defaults({
    encoding: null
});
var Promise = require('bluebird');
// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
   console.log('%s listening to %s', server.name, server.url); 
});
  
// Create chat connector for communicating with the Bot Framework Service
var connector = new builder.ChatConnector({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword,
    openIdMetadata: process.env.BotOpenIdMetadata 
});

// Listen for messages from users 
server.post('/api/messages', connector.listen());

/*----------------------------------------------------------------------------------------
* Bot Storage: This is a great spot to register the private state storage for your bot. 
* We provide adapters for Azure Table, CosmosDb, SQL Azure, or you can implement your own!
* For samples and documentation, see: https://github.com/Microsoft/BotBuilder-Azure
* ---------------------------------------------------------------------------------------- */

var tableName = 'botdata';
var azureTableClient = new botbuilder_azure.AzureTableClient(tableName, process.env['AzureWebJobsStorage']);
var tableStorage = new botbuilder_azure.AzureBotStorage({ gzipData: false }, azureTableClient);

// Create your bot with a function to receive messages from the user
var bot = new builder.UniversalBot(connector);


//bot.set('storage', tableStorage);   // Activate in  Production

// Make sure you add code to validate these fields
var luisAppId = process.env.LuisAppId;
var luisAPIKey = process.env.LuisAPIKey;
var luisAPIHostName = process.env.LuisAPIHostName || 'westus.api.cognitive.microsoft.com';

//const LuisModelUrl = 'https://' + luisAPIHostName + '/luis/v1/application?id=' + luisAppId + '&subscription-key=' + luisAPIKey;                  // Activate in  Production
const LuisModelUrl = 'https://westus.api.cognitive.microsoft.com/luis/v2.0/apps/9261e399-d445-4d44-b7f4-a0d79ec75c1f?subscription-key=9275fe239b584be3a66479224467dd5f&verbose=true&timezoneOffset=0';


//SQL Details
// Create connection to database
var config = 
   {
     userName: 'stanley1134', // update me
     password: 'Asdf@123', // update me
     server: 'stanazure.database.windows.net', // update me
     options: 
        {
           database: 'ChatBotSql' //update me
           , encrypt: true
        }
   }


   
  

   //User Authentication
   var settings = {
    url: "https://microexcel1.sharepoint.com/sites/appdev/hrdemo/",
    username: "nutan.sharma@microexcel.com",
    password: "meidYi9r9f6@7"
}
csomapi.setLoaderOptions({
    url: settings.url
});
var authCtx = new AuthenticationContext(settings.url);

// Main dialog with LUIS
var recognizer = new builder.LuisRecognizer(LuisModelUrl);
var intents = new builder.IntentDialog({ recognizers: [recognizer] })

.matches('GetWeather', (session) => {
    session.beginDialog('/GetWeather');
})

.matches('Greeting', (session) => {
    session.beginDialog('/Greetings');
})
.matches('Help', (session) => {
    session.beginDialog('/Help');
})
.matches('None', (session) => {
    session.beginDialog('/None');
})
.matches('Cancel', (session) => {
    session.beginDialog('/Confirm');
})
.matches('CreateTicket', (session) => {
  //  session.beginDialog('/CreateTicket');
})
.matches('LeaveRequest', (session) => {
    session.beginDialog('/CreateLeaveRequest');
})
.matches('ExpenseClaim', (session) => {
   // session.beginDialog('/CreateExpenseRequest');
})
.onDefault((session) => {
    session.beginDialog('/None');
});

bot.dialog('/', intents);   

bot.on('conversationUpdate', function(message) {
    if (message.membersAdded) {
        message.membersAdded.forEach(function(identity) {
            if (identity.id === message.address.bot.id) {
                var reply = new builder.Message()
                    .address(message.address)
                    .text("Welcome to Super Chat !" );
                bot.send(reply);
            }
        });
    }
});

bot.dialog('/Greetings', [
    function(session) {
        session.send('Hello %s, How can i help you ?', session.message.user.name).endDialog();
    }
]);
bot.dialog('/hotels-search',  [
    function(session) {
        session.send('Hello %s, How can i help you ?', session.message.user.name).endDialog();
    }
]);

bot.dialog('/Cancel', [
    function(session) {
        session.endConversation("Exiting the Process.Thanks for contacting. Have a wonderful day !");
        delete session.userData;
    }
]);
bot.dialog('/None', [
    function(session) {
        session.send("Sorry, I did not understand this '%s'", session.message.text).endDialog();
    }
]);
bot.dialog('/Confirm', [
    function(session) {
        builder.Prompts.confirm(session, "Do you wish to continue? (yes/no)");
    },
    function(session, results) {
        if (results.response == false) {
            session.beginDialog('/Cancel');
        } else {
            session.beginDialog('/Greetings');
        }
    }
]);

bot.dialog('/Help', function(session) {     
    var connection = new Connection(config);
    connection.on('connect', function(err) {
        if (err) {
            session.send('ERROR in fetching processes. Kindly contact the administrator.').endDialog();
            return;
        } else {
            var sqlQuery = 'SELECT [Process] FROM [dbo].[PersonalDetails]';
            
            console.log('Reading rows from the Table...');

            GetData(connection, sqlQuery, session, function(error, rowCount, results) {
                console.log("Processes Count: " + rowCount);
                if (rowCount > 0) {
                    var processes = '';
                    results.forEach(function(data) {
                        if (processes == "")
                            processes = data.Process;
                        else
                            processes += "|" + data.Process;
                    });
                    processes += "|Exit"
                    console.log("Processes List:" + processes)
                    session.privateConversationData['Processes'] = processes;
                    session.beginDialog('/ProcessInteraction');
                  
                
                   




                } else {
                    session.send("No results found for '%s'", session.message.text).endDialog();
                }

            });
          
        }
       
    });
});



//Process Interaction
bot.dialog('/ProcessInteraction', [
    function(session, args) {
        if (session.privateConversationData['Processes'] != "") {
            session.userData.ProcessName = args || {};
            builder.Prompts.choice(session, "Choose area of Interest:", session.privateConversationData['Processes'], {
                listStyle: builder.ListStyle.button
            });
        } else {
            session.send("No Process Found.");
            session.endDialog();
        }
    },
    function(session, results) {
        session.userData.ProcessName = results.response.entity;
        session.send("You chose:" + session.userData.ProcessName);

        if (session.userData.ProcessName == 'Ticket Agent') {
            session.beginDialog('/CreateTicket');
        } else if (session.userData.ProcessName == 'Leave Request') {
            session.beginDialog('/CreateLeaveRequest');
        } else if (session.userData.ProcessName == 'Expense Reimbursement') {
            session.beginDialog('/CreateExpenseRequest');
        } else if (session.userData.ProcessName == 'FAQ') {
            session.beginDialog('/FAQInputs');
            session.endDialog();
        } else if (session.userData.ProcessName == 'Exit') {
            session.beginDialog('/Confirm');
        } else {
            session.beginDialog('/Cancel');
        }
    },
    function(session, results) {
        if(session.userData.ProcessName != 'FAQ')
        session.beginDialog('/Confirm');
        else
         session.endDialog();
    }
]);

//Functions
function GetData(connection, sqlQuery, session, callback) {

    var results = [];

    console.log("QUERY: " + sqlQuery);
    var request = new Request(sqlQuery, function(error, rowCount, rows) {
        if (error) {
            return callback(error);
        }
        callback(null, rowCount, results);
    });

    request.on('row', function(columns) {
        var row = {};
        columns.forEach(function(column) {
            row[column.metadata.colName] = column.value.trim();
        });
        results.push(row);
    });
    connection.execSql(request);
}


//Sharepoint Functions
bot.dialog('/CreateLeaveRequest', [

    function(session, args) {
        session.sendTyping();
        session.userData.LeaveDetails = args || {};
        builder.Prompts.text(session, "Subject:");
    },
    function(session, results) {
        session.userData.LeaveDetails.Title = results.response;
        builder.Prompts.choice(session, "Leave Type:", "Vacation|Sick Leave|Floating Holiday|Jury Duty|Other", {
            listStyle: builder.ListStyle.button
        });
    },
    function(session, results) {
        session.userData.LeaveDetails.LeaveType = results.response.entity;
        builder.Prompts.time(session, "Start Date:");
    },
    function(session, results) {
        session.userData.LeaveDetails.StartDate = builder.EntityRecognizer.resolveTime([results.response]);;
        builder.Prompts.time(session, "End Date:");
    },
    function(session, results) {
        session.userData.LeaveDetails.EndDate = builder.EntityRecognizer.resolveTime([results.response]);;
        builder.Prompts.text(session, "Reason for Leave:");
    },
    function(session, results) {
        session.userData.LeaveDetails.LeaveReason = results.response;
        builder.Prompts.text(session, "Manager:");
    },
    function(session, results) {
        session.userData.LeaveDetails.Manager = results.response;
        console.log(session.userData.LeaveDetails);
        CreateLeaveRequestInSharePoint(session.userData.LeaveDetails, function(itemTitle, itemUrl, isSuccess) {
            if (isSuccess) {
                var card = new builder.HeroCard(session)
                    .title("You request is successfully submitted.")
                    .buttons([
                        builder.CardAction.openUrl(session, itemUrl, "View")
                    ]);
                var msg = new builder.Message(session).addAttachment(card);
                session.send(msg);
            } else
                session.send("Error in submitting request.Kindly contact System Administrator.");

            session.beginDialog('/Confirm');
        });
    }
]);


//Create Leave in Sharepoint
function CreateLeaveRequestInSharePoint(LeaveDetails, fn) {
    authCtx.acquireTokenForUser(settings.username, settings.password, function(err, data) {
        var ctx = new SP.ClientContext(settings.url);
        authCtx.setAuthenticationCookie(ctx);
        var web = ctx.get_web();
        ctx.load(web);
        ctx.executeQueryAsync(function() {
                var ctx = web.get_context();
                var list = web.get_lists().getByTitle("PTO Requests");
                var creationInfo = new SP.ListItemCreationInformation();

                var listItem = list.addItem(creationInfo);
                listItem.set_item('Title', LeaveDetails.Title);
                listItem.set_item('LeaveType', LeaveDetails.LeaveType);
                listItem.set_item('StartDate', LeaveDetails.StartDate);
                listItem.set_item('EndDate', LeaveDetails.EndDate);
                listItem.set_item('LeaveReason', LeaveDetails.LeaveReason);

                var manager = SP.FieldUserValue.fromUser(LeaveDetails.Manager);
                if (manager)
                    listItem.set_item('Manager', manager);

                listItem.update();

                ctx.load(listItem);
                ctx.executeQueryAsync(function() {
                        fn(LeaveDetails.Title, settings.url + "Lists/PTORequest/dispform.aspx?ID=" + listItem.get_id(), true);
                    },
                    function(sender, args) {
                        console.log('An error occured: ' + args.get_message());
                        fn(args.get_message());
                    });

            },
            function(sender, args) {
                console.log('An error occured: ' + args.get_message());
                fn(args.get_message());
            });
    });
}



//Weather
bot.dialog('/GetWeather', [
   
    function(session, args) {
        var city = "FishKill";
        var searchtext = "select * from weather.forecast where woeid in (select woeid from geo.places(1) where text='"+city+", ak')"
        //change city variable dynamically as required
    
        getJSON("https://query.yahooapis.com/v1/public/yql?q=" + searchtext + "&format=json", function(error, response){
     
            var data = response.query.results;
            var msg = new builder.Message(session)
            .addAttachment(weathercardObj.weatherCard("55"));
        session.send(msg);
        })
    
    }

]);