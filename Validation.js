var builder = require('botbuilder');
exports.ValidateEmail = function (session,mailObj) {
    
  var IsValidated = true;
  if(mailObj.emailTo == "")
  {
    IsValidated = false;
    
    var card = new builder.HeroCard(session)
    .title("Error")
    .text("Enter valid Email Address")

    var msg = new builder.Message(session).addAttachment(card);
    session.send(msg);
  }
  else if(!validateEmail(mailObj.emailTo))
  {
    IsValidated = false;
    var card = new builder.HeroCard(session)
    .title("Error")
    .text("Enter valid Email Address")

    var msg = new builder.Message(session).addAttachment(card);
    session.send(msg);
  }
  
return IsValidated;
}

exports.ValidateLeaveRequestForm = function (session,LeaveRequestObj) {
    
    var IsValidated = true;
    if(LeaveRequestObj.LeaveSubject == "")
    {
      IsValidated = false;
      
      var card = new builder.HeroCard(session)
      .title("Error")
      .text("Enter Subject")
  
      var msg = new builder.Message(session).addAttachment(card);
      session.send(msg);  
    }
    else if(LeaveRequestObj.LeaveType == "")
    {
      IsValidated = false;
      
      var card = new builder.HeroCard(session)
      .title("Error")
      .text("Select Leave Type")
  
      var msg = new builder.Message(session).addAttachment(card);
      session.send(msg);  
    }
    else if(LeaveRequestObj.LeaveStartDate == "")
    {
      IsValidated = false;
      
      var card = new builder.HeroCard(session)
      .title("Error")
      .text("Select Start Date")
  
      var msg = new builder.Message(session).addAttachment(card);
      session.send(msg);  
    }
    
    else if(LeaveRequestObj.LeaveEndDate == "")
    {
      IsValidated = false;
      
      var card = new builder.HeroCard(session)
      .title("Error")
      .text("Select End Date")
  
      var msg = new builder.Message(session).addAttachment(card);
      session.send(msg);  
    
    }
    else if(LeaveRequestObj.LeaveReason == "")
    {
      IsValidated = false;
      
      var card = new builder.HeroCard(session)
      .title("Error")
      .text("Enter Reason")
  
      var msg = new builder.Message(session).addAttachment(card);
      session.send(msg);  
    
    }
    else if(LeaveRequestObj.Leavemanager == "")
    {
      IsValidated = false;
      
      var card = new builder.HeroCard(session)
      .title("Error")
      .text("Enter Manager")
  
      var msg = new builder.Message(session).addAttachment(card);
      session.send(msg);  
    
    }
    else if(LeaveRequestObj.LeaveStartDate > LeaveRequestObj.LeaveEndDate)
    {
      IsValidated = false;      
      var card = new builder.HeroCard(session)
      .title("Error")
      .text("End Date must be greater than Start Date")
  
      var msg = new builder.Message(session).addAttachment(card);
      session.send(msg); 
    
    }
   
    
  return IsValidated;
  }

  exports.TicketRequestFormValidation = function (session,tktObj) {
    
    var IsValidated = true;
    if(tktObj.tktSubject == "")
    {
      IsValidated = false;
      
      var card = new builder.HeroCard(session)
      .title("Error")
      .text("Enter Subject")
  
      var msg = new builder.Message(session).addAttachment(card);
      session.send(msg);
  
    }
    else if(tktObj.tktDescription == "")
    {
      IsValidated = false;
      
      var card = new builder.HeroCard(session)
      .title("Error")
      .text("Enter Description")
  
      var msg = new builder.Message(session).addAttachment(card);
      session.send(msg);
  
    }
    else if(tktObj.tktCategory == "")
    {
      IsValidated = false;
      
      var card = new builder.HeroCard(session)
      .title("Error")
      .text("Select Category")
  
      var msg = new builder.Message(session).addAttachment(card);
      session.send(msg);
  
    }
    
    else if(tktObj.tktPriority == "")
    {
      IsValidated = false;
      
      var card = new builder.HeroCard(session)
      .title("Error")
      .text("Select Priority")
  
      var msg = new builder.Message(session).addAttachment(card);
      session.send(msg);
  
    
    }
   
    
   
    
  return IsValidated;
  }

  exports.ExpenseFormValidation = function (session,expObj) {
    
    var IsValidated = true;
    if(expObj.expName == "")
    {
      IsValidated = false;
      
      var card = new builder.HeroCard(session)
      .title("Error")
      .text("Enter Expense Name")
  
      var msg = new builder.Message(session).addAttachment(card);
      session.send(msg);
  
    }
    else if(expObj.expType == "")
    {
      IsValidated = false;
      
      var card = new builder.HeroCard(session)
      .title("Error")
      .text("Select Expense Type")
  
      var msg = new builder.Message(session).addAttachment(card);
      session.send(msg);
  
    }
    else if(expObj.ExpCategory == "")
    {
      IsValidated = false;
      
      var card = new builder.HeroCard(session)
      .title("Error")
      .text("Select Category")
  
      var msg = new builder.Message(session).addAttachment(card);
      session.send(msg);
  
    }
    
    else if(expObj.expAmount == "")
    {
      IsValidated = false;
      
      var card = new builder.HeroCard(session)
      .title("Error")
      .text("Enter Amount")
  
      var msg = new builder.Message(session).addAttachment(card);
      session.send(msg);
  
    
    }
   
    
   
    
  return IsValidated;
  }
function validateEmail(email) {
    var re = /^(([^<>()[\]\\.,;:\s@\"]+(\.[^<>()[\]\\.,;:\s@\"]+)*)|(\".+\"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    return re.test(email);
  }
  