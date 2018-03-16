
exports.emailCard = function (statusList,status) {
    
  
    var emailCardObj = {
        'contentType': 'application/vnd.microsoft.card.adaptive',
        'content': {
          "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
          "type": "AdaptiveCard",
          "version": "1.0",
          "body": [
            {
                "type": "ColumnSet",
                "columns": [
                  {
                    "type": "Column",
                    "width": 2,
                    "items": [
                      {
                        "type": "TextBlock",
                        "text": "Send Mail",
                        "weight": "bolder",
                        "size": "medium"
                      },
                     
                      {
                        "type": "TextBlock",
                        "text": "",
                        "isSubtle": true,
                        "wrap": true,
                        "size": "small"
                      },
                      {
                        "type": "TextBlock",
                        "text": "To *",
                        "wrap": true,
                        "Required" :"Yes"
                      },
                      {
                        "type": "Input.Text",
                        "placeholder":"youremail@example.com",
                        "style": "email",
                        "maxLength": 0,
                        "id": "emailTo",
                        "IsRequired":true
                      },
                      {
                        "type": "TextBlock",
                        "text": "Subject",
                        "wrap": true
                      },
                      {
                        "type": "Input.Text",
                        "id": "emailSubject",
                        
                      },
                      
                      {
                        "type": "TextBlock",
                        "text": "Body"
                      },
                      {
                        "type": "Input.Text",
                        "id": "emailBody",
                        "isMultiline": true,
                        "style": "text",
                        "maxLength": 0,
                        "width":"auto"
                        
                      },
                     
                      
                    ]
                  },
                  {
                    "type": "Column",
                    "width": 1,
                    "items": [
                      {
                        "type": "Image",
                        "url": "",
                        "size": "auto"
                      }
                    ]
                  }
                ]
              }
          ],
          "actions": [
            {
                'type': 'Action.Submit',
                'title': 'Submit',
                
                'data': {
                    'type': 'emailSubmit'
                }
            }
        ]
             
  }
}
return emailCardObj;
}