
exports.ticketInputForm = function () {
    
  
    var tktinputformCard = {
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
                      "text": "Create Ticket Form",
                      "weight": "bolder",
                      "size": "medium"
                    },
                    {
                      "type": "TextBlock",
                      "text": "Create Ticket Description",
                      "isSubtle": true,
                      "wrap": true
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
                      "text": "Ticket Subject",
                      "wrap": true
                    },
                    {
                      "type": "Input.Text",
                      "id": "tktSubject",
                      "placeholder": "Subject"
                    },
                    {
                        "type": "TextBlock",
                        "text": "Enter Description"
                      },
                      {
                        "type": "Input.Text",
                        "id": "tktDescription",
                        "isMultiline": true,
                        "style": "text",
                        "maxLength": 0,
                        
                      },
                    {
                      "type": "TextBlock",
                      "text": "Category",
                      "wrap": true
                    },
                    {
                      "type": "Input.ChoiceSet",
                      "id": "tktCategory",
                      "style": "compact",
                      "value": "1",
                      
                      "choices": [
                        
                          {
                            "title": "HR",
                            "value": "HR"
                          },
                          {
                            "title": "Facilities",
                            "value": "Facilities"
                          },
                          {
                              "title": "Operations",
                              "value": "Operations"
                            },
                          

                        ]
                    },
                    {
                        "type": "TextBlock",
                        "text": "Priority",
                        "wrap": true
                      },
                      {
                        "type": "Input.ChoiceSet",
                        "id": "tktPriority",
                        "style": "compact",
                        "value": "1",
                        
                        "choices": [
                            {
                              "title": "High",
                              "value": "High"
                            },
                            {
                              "title": "Medium",
                              "value": "Medium"
                            },
                            {
                              "title": "Low",
                              "value": "Low"
                            },
                                                       
  
                          ]
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
                  'speak': '<s>Search</s>',
                  'data': {
                      'type': 'ticketFormSubmit'
                  }
              }
          ]


    
  }
}
return tktinputformCard;
}