/*-----------------------------------------------------------------------------
A simple echo bot for the Microsoft Bot Framework. 
-----------------------------------------------------------------------------*/

var SploginUserName = "nutan.sharma@microexcel.com";
var SploginPassword = "meidYi9r9f6@7";

var dt = require('./data.js');
var $ = require('jQuery');
var getJSON = require('get-json');
var get = require('get');
var feed = require("feed-read-parser");
var spauth = require('node-sp-auth');
var nodemailer = require('nodemailer');

var cardObj = require('./card.js');
var weathercardObj = require('./weathercard.js');
var InputcardObj = require('./inputform.js');
var ticketcardObj = require('./ticketInputForm.js');
var expcardObj = require('./expenseForm.js');
var statusCardObj = require('./StatusForm.js');
var EmailCardObj = require('./emailCard.js');
var MailValidationObj = require('./Validation.js');



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

   //Email Configration
   var transporter = nodemailer.createTransport({
    service: 'sendgrid',
    auth: {
      user: 'azure_5ab59171ecc0e95d22c6ed723b88bd38@azure.com',
      pass: 'Asdf@123',
      port:587
    }
  });
   
  var FromEmail = 'ChatBot<azure_5ab59171ecc0e95d22c6ed723b88bd38@azure.com>';
  

   //User Authentication
   var settings = {
    url: "https://microexcel1.sharepoint.com/sites/appdev/hrdemo/",
    username: SploginUserName,
    password: SploginPassword
}
csomapi.setLoaderOptions({
    url: settings.url
});
var authCtx = new AuthenticationContext(settings.url);

// Main dialog with LUIS
var recognizer = new builder.LuisRecognizer(LuisModelUrl);
var intents = new builder.IntentDialog({ recognizers: [recognizer] })

.matches('GetWeather', (session, args, next) => {
    session.beginDialog('/GetWeather',args);
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
    session.beginDialog('/CreateTicket');
})
.matches('LeaveRequest', (session) => {
    session.beginDialog('/CreateLeaveRequestInputForm');
})
.matches('ExpenseClaim', (session) => {
    session.beginDialog('/CreateExpenseRequest');
})
.matches('GetStatus', (session, args, next) => {
     session.beginDialog('/StatusCheckSPList',args);
 })
.matches('GetFaq', (session) => {
    session.beginDialog('/GetCommunitySites');
 })
 .matches('SendMail', (session) => {
    session.beginDialog('/SendEmail');
 })
.onDefault((session,args) => {

    if (session.message && session.message.value) {
        // A Card's Submit Action obj was received
        processSubmitAction(session,args, session.message.value);
        return;
    }
    session.beginDialog('/None');
});



bot.dialog('/', intents);   

bot.on('conversationUpdate', function(message) {
    if (message.membersAdded) {
        message.membersAdded.forEach(function(identity) {
            if (identity.id === message.address.bot.id) {
                var reply = new builder.Message()
                    .address(message.address)
                    .text("Welcome to HR Processes !" );
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
bot.dialog('/blank', [
    function(session) {
        session.send('').endDialog();
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

bot.dialog('/SubmitForm', [
    function(session) {
        session.send("U r in submit form").endDialog();
    }
]);
bot.dialog('/Confirm', [
    function(session) {
        builder.Prompts.confirm(session, "Do you wish to continue? (yes/no)");
    },
    function(session, results) {
        if (results.response == false) {
            session.beginDialog('/blank');
        } else {
            session.beginDialog('/Greetings');
        }
    }
]);




bot.dialog('/Help', [ 
    
    function(session, args) {
      
            session.userData.HelpProcess = args || {};
            builder.Prompts.choice(session, "Choose area of Interest:", "FAQ|Processes", {
                listStyle: builder.ListStyle.button
            });
        
    },

    function(session, results) {

        if(results.response.entity == "FAQ")
        {
        session.userData.HelpProcess = results.response.entity;
       
        session.send("Enter your query");
        session.endDialog();
        }
        else if(results.response.entity == "Processes")
        {
            session.endDialog();
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
        }
    
},
    



]

);



//Process Interaction
bot.dialog('/ProcessInteraction', [
    function(session, args) {
        if (session.privateConversationData['Processes'] != "") {
            session.userData.ProcessName = args || {};
            builder.Prompts.choice(session, "Choose a process:", session.privateConversationData['Processes'], {
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
            session.beginDialog('/CreateLeaveRequestInputForm');
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
    /*function(session, results) {
        if(session.userData.ProcessName != 'FAQ')
        session.beginDialog('/Confirm');
        else
         session.endDialog();
    }*/
]);

//Process Interaction
bot.dialog('/StatusCheckSPList', [
    function(session, args) {

       
        if(args.intent == "GetStatus")
        {
            if(args.entities.length >0)
            {
            var subtopicEntity = builder.EntityRecognizer.findEntity(args.entities, 'trackingID');
            var indx = subtopicEntity.entity.toString().indexOf("pto") ;

            if(subtopicEntity.entity.toString().indexOf("pto") > -1) {
                GetStatusFromSharepointList(session,'PTO Requests',subtopicEntity.entity,'TrackingID','leave');
            }
            else if(subtopicEntity.entity.toString().indexOf("req") > -1) {
                GetStatusFromSharepointList(session,'Tickets',subtopicEntity.entity,'RequestID','ticket');
            }

            

            session.beginDialog('/blank');
        }
    }

    },
   
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









//Weather
bot.dialog('/GetWeather', [
   
    function(session,args,next) {
        if(args.intent == "GetWeather")
        {
            if(args.entities.length >0)
            {
            var subtopicEntity = builder.EntityRecognizer.findEntity(args.entities, 'City');
                   // var city = "FishKill";
        var searchtext = "select * from weather.forecast where woeid in (select woeid from geo.places(1) where text='"+subtopicEntity.entity+", ak')"
        //change city variable dynamically as required
    
        getJSON("https://query.yahooapis.com/v1/public/yql?q=" + searchtext + "&format=json", function(error, response){
     
            var data = response.query.results;
            var msg = new builder.Message(session)
            .addAttachment(weathercardObj.weatherCard(data.channel.item.condition.code,data.channel.item.condition.temp,data.channel.item.forecast[0].high,data.channel.item.forecast[0].low,data.channel.location.city,data.channel.location.region,data.channel.item.condition.text));
        session.send(msg);
        })


        }
    }
    session.beginDialog('/blank');
    }
    
]);



//Community Sites
bot.dialog('/GetCommunitySites', [
   
    function(session,args,next) {
       
        session.send("Searching...");
       
        var communitySiteUrl = 'https://support.office.com';
        var SharepointSiteUrl = 'https://microexcel1.sharepoint.com'
        
        feed("https://www.bing.com/search?q='"+session.message.text+"'+site:https://support.office.com&format=rss&count=2", function(err, articles) {
            if(articles.length > 0)
            {

                var attachments = [];
                var msg = new builder.Message(session);
                msg.attachmentLayout(builder.AttachmentLayout.carousel);

                articles.forEach(function(data) {
                    var card = new builder.HeroCard(session)    
                        .title(data.title)
                        .text(data.content)
                        .images([builder.CardImage.create(session, 'https://img-prod-cms-rt-microsoft-com.akamaized.net/cms/api/am/imageFileData/RE1Mu3b?ver=5c31')])
                        .buttons([
                            builder.CardAction.openUrl(session, data.link, "Get Started")
                        ])
                    attachments.push(card);
                });

                msg.attachments(attachments);
                //session.send(msg);
                SearchSharepoint(session,attachments);
            }
            else{
                session.send("No results found");
            }
            
           });
    session.beginDialog('/blank');
    }
    
]);


//Search
function SearchSharepoint(session,attachments) {

    spauth
    .getAuth('https://microexcel1.sharepoint.com/sites/appdev/hrdemo/', {
      username: SploginUserName,
      password: SploginPassword
    })
    .then(data => {
      var headers = data.headers;
      headers['Accept'] = 'application/json;odata=verbose';
  
      requestpromise.get({
        url: "https://microexcel1.sharepoint.com/_api/search/query?querytext='"+session.message.text+"+path:https://microexcel1.sharepoint.com/sites/appdev/hrdemo'",
        headers: headers,
        json: true
      }).then(response => {
        //var fields = response.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results.Cells;
        var query = response.d.query; 
      var resultsCount = query.PrimaryQueryResult.RelevantResults.RowCount;

      //var attachments = [];
      var msg = new builder.Message(session);
      msg.attachmentLayout(builder.AttachmentLayout.carousel);

      for(var i = 0; i < resultsCount;i++) {
          var row = query.PrimaryQueryResult.RelevantResults.Table.Rows.results[i];
          var Title = row.Cells.results[3].Value;
          var pathurl = row.Cells.results[6].Value;
          var Description = row.Cells.results[11].Value;
          Description =  Description.replace(/<[^>]+>/g, '');
        
          var card = new builder.HeroCard(session)    
            .title(Title)
            .text(Description)
            .images([builder.CardImage.create(session, 'http://www.microexcel.com/wp-content/uploads/2018/01/Microexcel-Logo.png')])
            .buttons([
                builder.CardAction.openUrl(session, pathurl, "Get Started")
            ])
                attachments.push(card);
      }   
      msg.attachments(attachments);
      session.send(msg);
      session.beginDialog('/blank');
      });
    });
    session.beginDialog('/blank');
}


//Get Status From Sharepoint
function GetStatusFromSharepointList(session,listName,trackingID,columnName,StatusType) {

    spauth
    .getAuth('https://microexcel1.sharepoint.com/sites/appdev/hrdemo/', {
      username: SploginUserName,
      password: SploginPassword
    })
    .then(data => {
      var headers = data.headers;
      headers['Accept'] = 'application/json;odata=verbose';
  
      requestpromise.get({
        url: "https://microexcel1.sharepoint.com/sites/appdev/hrdemo/_api/web/lists/getbytitle('"+listName+"')/items?$filter= "+columnName+" eq '"+trackingID+"'",
        headers: headers,
        json: true
      }).then(response => {
        //var fields = response.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results.Cells;
        if(response.d.results.length >0)
        {
        var status = response.d.results[0].Status;         
        var msg = new builder.Message(session)
        .addAttachment(statusCardObj.StatusForm("Ticket Status : "+trackingID+"",status));
         session.send(msg);
  
        }
        else{
            session.send("Unable to locate Tracking ID");
        }
      
      
      session.beginDialog('/blank');
      
      session.beginDialog('/blank');
      });
    });
    session.beginDialog('/blank');
}



//Sharepoint Functions Input Form
bot.dialog('/CreateLeaveRequestInputForm', [

   
    function(session,args,next) {
        
            
        var msg = new builder.Message(session)
            .addAttachment(InputcardObj.inputform());
             session.send(msg);
    
             session.endDialog();
          
            
    }
   
   
   

]);
//Create Ticket
bot.dialog('/CreateTicket', [

    function(session,args,next) {
                    
        var msg = new builder.Message(session)
            .addAttachment(ticketcardObj.ticketInputForm());
             session.send(msg);    
             session.endDialog();          
            
    }

]);


//Create Expense Request
bot.dialog('/CreateExpenseRequest', [

    function(session, results) {
        builder.Prompts.attachment(session, "Attach Copy of Receipt:");
    },
    
    function(session,args,next) {
        session.userData.ClaimDetails = args || {};
        session.userData.ClaimDetails.Receipt = session.message.attachments[0].name;
        session.userData.ClaimDetails.ContentUrl = session.message.attachments[0].contentUrl;
        

        var msg = new builder.Message(session)
            .addAttachment(expcardObj.expenseForm());
             session.send(msg);    
             session.endDialog();          
            
    }


]);

// Send Mail
bot.dialog('/SendEmail', [ 
    
    function(session, args) {


      
        var msg = new builder.Message(session)
            .addAttachment(EmailCardObj.emailCard());
             session.send(msg);
    
             session.endDialog();


        
    },

  



]

);


//Create Leave Request in Sharepoint
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
                var PtoGuid = ptoguid();
                var listItem = list.addItem(creationInfo);
                listItem.set_item('Title', LeaveDetails.Title);
                listItem.set_item('LeaveType', LeaveDetails.LeaveType);
                listItem.set_item('StartDate', LeaveDetails.StartDate);
                listItem.set_item('EndDate', LeaveDetails.EndDate);
                listItem.set_item('LeaveReason', LeaveDetails.LeaveReason);
                listItem.set_item('TrackingID', PtoGuid);

                var manager = SP.FieldUserValue.fromUser(LeaveDetails.Manager);
                if (manager)
                    listItem.set_item('Manager', manager);

                listItem.update();

                ctx.load(listItem);
                ctx.executeQueryAsync(function() {
                        fn(LeaveDetails.Title,PtoGuid, settings.url + "Lists/PTORequest/dispform.aspx?ID=" + listItem.get_id(), true);
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
//Create expense report in sharepoint
function CreateExpenseClaimInSharePoint(session, attachmentDetails, ClaimDetails, fn) {
    authCtx.acquireTokenForUser(settings.username, settings.password, function(err, data) {
        var ctx = new SP.ClientContext(settings.url);
        authCtx.setAuthenticationCookie(ctx);
        var web = ctx.get_web();
        ctx.load(web);
        ctx.executeQueryAsync(function() {
                var ctx = web.get_context();
                var list = web.get_lists().getByTitle("Expense Claims");
                var creationInfo = new SP.ListItemCreationInformation();
                var listItem = list.addItem(creationInfo);

                listItem.set_item('Title', ClaimDetails.Title);
                listItem.set_item('ExpenseType', ClaimDetails.ExpenseType);
                listItem.set_item('ExpenseCategory', ClaimDetails.ExpenseCategory);
                listItem.set_item('Customer', ClaimDetails.Customer);
                listItem.set_item('ExpenseAmount', ClaimDetails.ExpenseAmount);
                listItem.update();

                ctx.load(listItem);
                ctx.executeQueryAsync(function() {
                        var attachmentsFolder = web.getFolderByServerRelativeUrl('Lists/ExpenseClaims/Attachments/');
                        var attachmentRootFolderUrl=  web.get_serverRelativeUrl()+ 'Lists/ExpenseClaims/Attachments'; 
                        attachmentsFolder = attachmentsFolder.get_folders().add('_' + listItem.get_id());
                        attachmentsFolder.moveTo(attachmentRootFolderUrl + '/' + listItem.get_id()); 
                        ctx.load(attachmentsFolder);  
                        ctx.executeQueryAsync(function() {                        
                            var attachmentFolder = web.getFolderByServerRelativeUrl('Lists/ExpenseClaims/Attachments/' + listItem.get_id());
                            var createInfo = new SP.FileCreationInformation();
                            createInfo.set_content(attachmentDetails);
                            createInfo.set_url(ClaimDetails.Receipt);
                            attachmentFiles = attachmentFolder.get_files().add(createInfo);
                            ctx.load(list);
                            ctx.load(attachmentFiles);
                            ctx.executeQueryAsync(
                                function() {
                                    fn(ClaimDetails.Title, settings.url + "Lists/ExpenseClaims/dispform.aspx?ID=" + listItem.get_id(), true);
                                },
                                function(sender, args) {
                                    console.log('Adding Attachment. An error occured: ' + args.get_message());
                                    fn(args.get_message());
                                });
                            },
                                function(sender, args) {
                                    console.log('Adding Attachment. An error occured: ' + args.get_message());
                                    fn(args.get_message());
                           });
                    },
                    function(sender, args) {
                        console.log('Adding Item:An error occured: ' + args.get_message());
                        fn(args.get_message());
                    });

            },
            function(sender, args) {
                console.log('Authentication. An error occured: ' + args.get_message());
                fn(args.get_message());
            });
    });
}

//Create ticket in sharepoint
function CreateTicketInSharePoint(TicketDetails, fn) {
    authCtx.acquireTokenForUser(settings.username, settings.password, function(err, data) {
        var ctx = new SP.ClientContext(settings.url);
        authCtx.setAuthenticationCookie(ctx);
        var web = ctx.get_web();
        ctx.load(web);
        ctx.executeQueryAsync(function() {
                var ctx = web.get_context();
                var list = web.get_lists().getByTitle("Tickets");
                var creationInfo = new SP.ListItemCreationInformation();
                var listItem = list.addItem(creationInfo);

                var requestId = guid();
                listItem.set_item('RequestID', requestId);
                listItem.set_item('Title', TicketDetails.Title);
                listItem.set_item('Description', TicketDetails.Description);
                listItem.set_item('Category', TicketDetails.Category);
                listItem.set_item('Priority', TicketDetails.Priority);
                listItem.update();

                ctx.load(listItem);
                ctx.executeQueryAsync(function() {
                        fn(requestId, settings.url + "Lists/Tickets/dispform.aspx?ID=" + listItem.get_id(), true);
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

//Send Mail
function SendEmailFromChat(EmailDetails, fn) {
    
              
      var mailOptions = {
        from: FromEmail ,
        to: EmailDetails.ToAddress,
        subject: EmailDetails.Subject,
        text: EmailDetails.Body
      };
      
      transporter.sendMail(mailOptions, function(error, info){
        if (error) {
          console.log(error);
        } else {
            fn(true);          
        }
      });
    
}

function processSubmitAction(session,args, value) {
    var defaultErrorMessage = 'Please complete all the search parameters';
    switch (value.type) {
        case 'inputFormSubmit':
            if(MailValidationObj.ValidateLeaveRequestForm(session,value))
            {
                session.userData.LeaveDetails = args || {};
                session.userData.LeaveDetails.Title = value.LeaveSubject;
                session.userData.LeaveDetails.LeaveType = value.LeaveType;
                session.userData.LeaveDetails.StartDate = value.LeaveStartDate;
                session.userData.LeaveDetails.EndDate = value.LeaveEndDate;
                session.userData.LeaveDetails.LeaveReason = value.LeaveReason;
                session.userData.LeaveDetails.Manager = value.Leavemanager;
            
                CreateLeaveRequestInSharePoint(session.userData.LeaveDetails, function(itemTitle,ptoguid, itemUrl, isSuccess) {
                    if (isSuccess) {
                        var card = new builder.HeroCard(session)
                            .title("You request is successfully submitted.")
                            .text("Tracking ID "+ptoguid+"")
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
            break;
            case 'ticketFormSubmit':
                if(MailValidationObj.TicketRequestFormValidation(session,value))
                {
                    session.userData.TicketDetails = args || {};
                    session.userData.TicketDetails.Title  = value.tktSubject;
                    session.userData.TicketDetails.Description  = value.tktDescription;
                    session.userData.TicketDetails.Category  = value.tktCategory;
                    session.userData.TicketDetails.Priority = value.tktPriority;
                

                    CreateTicketInSharePoint(session.userData.TicketDetails, function(requestID, itemUrl, isSuccess) {
                        if (isSuccess) {
                            var card = new builder.HeroCard(session)
                                .title("You request is successfully submitted.")
                                .text("Tracking ID "+requestID+"")
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
            break;
        

            case 'expFormSubmit':
            if(MailValidationObj.ExpenseFormValidation(session,value))
            {
            session.userData.ClaimDetails.Title  = value.expName;
            session.userData.ClaimDetails.ExpenseType  = value.expType;
            session.userData.ClaimDetails.Customer  = value.expCustomerName;
            session.userData.ClaimDetails.ExpenseCategory = value.ExpCategory;
            session.userData.ClaimDetails.ExpenseAmount = value.expAmount;
           

            var fileDownload = checkRequiresToken(session.message) ? requestWithToken(session.userData.ClaimDetails.ContentUrl) : requestpromise(session.userData.ClaimDetails.ContentUrl);
            fileDownload.then(
                function(response) {
                    var imageBase64Sting = new Buffer(response, 'binary').toString('base64');
    
                    CreateExpenseClaimInSharePoint(session, imageBase64Sting, session.userData.ClaimDetails, function(itemTitle, itemUrl, isSuccess) {
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
                }).catch(function(err) {
                console.log('Error downloading attachment:' + err);
                // console.log('Error downloading attachment:', { statusCode: err.statusCode, message: err.response.statusMessage });
            });
        }

        
            break;
      
            case 'emailSubmit':
            session.userData.EmailDetails = args || {};
            session.userData.EmailDetails.FromAddress  = "";
            if(MailValidationObj.ValidateEmail(session,value))
            {
            session.userData.EmailDetails.ToAddress  = value.emailTo;
            session.userData.EmailDetails.Subject  = value.emailSubject;
            session.userData.EmailDetails.Body = value.emailBody;
            
            SendEmailFromChat(session.userData.EmailDetails, function(isSuccess) {
                if (isSuccess) {
                    var card = new builder.HeroCard(session)
                        .title("The mail has been sent successfully.")
                  
                    var msg = new builder.Message(session).addAttachment(card);
                    session.send(msg);
                } else
                    session.send("Error in submitting request.Kindly contact System Administrator.");
    
                session.beginDialog('/blank');
            });
        }
        



        
            break;

        default:
            // A form data was received, invalid or incomplete since the previous validation did not pass
            session.send(defaultErrorMessage);
    }
}
function guid() {
    function s4() {
        return Math.floor((1 + Math.random()) * 0x10000)
            .toString(16)
            .substring(1);
    }
    return "REQ" + s4();
}
function ptoguid() {
    function s4() {
        return Math.floor((1 + Math.random()) * 0x10000)
            .toString(16)
            .substring(1);
    }
    return "PTO" + s4();
}

// Request file with Authentication Header
var requestWithToken = function(url) {
    return obtainToken().then(function(token) {
        return requestpromise({
            url: url,
            headers: {
                'Authorization': 'Bearer ' + token,
                'Content-Type': 'application/octet-stream'
            }
        });
    });
};

// Promise for obtaining JWT Token (requested once)
var obtainToken = Promise.promisify(connector.getAccessToken.bind(connector));

var checkRequiresToken = function(message) {
    return message.source === 'skype' || message.source === 'msteams';
};