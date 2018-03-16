
exports.expenseForm = function () {
    
  
    var expenseFormCard = {
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
                      "text": "Expense Reimbursement Form",
                      "weight": "bolder",
                      "size": "medium"
                    },
                    {
                      "type": "TextBlock",
                      "text": "Expense Reimbursement Description",
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
                      "text": "Expense Name",
                      "wrap": true
                    },
                    {
                      "type": "Input.Text",
                      "id": "expName",
                      "placeholder": "Expense Name"
                    },
                    
                    {
                      "type": "TextBlock",
                      "text": "Expense Type",
                      "wrap": true
                    },
                    {
                      "type": "Input.ChoiceSet",
                      "id": "expType",
                      "style": "compact",
                      "value": "1",
                      
                      "choices": [
                          {
                            "title": "Internal",
                            "value": "Internal"
                          },
                          {
                            "title": "Customer",
                            "value": "Customer"
                          },
                          
                        ]
                    },
                    {
                        "type": "TextBlock",
                        "text": "Customer Name",
                        "wrap": true
                      },
                      {
                        "type": "Input.Text",
                        "id": "expCustomerName",
                        "placeholder": "Customer Name"
                      },
                    {
                        "type": "TextBlock",
                        "text": "Expense Category",
                        "wrap": true
                      },
                      {
                        "type": "Input.ChoiceSet",
                        "id": "ExpCategory",
                        "style": "compact",
                        "value": "1",
                        
                        "choices": [
                            {
                              "title": "Food",
                              "value": "Food"
                            },
                            {
                              "title": "Entertainment",
                              "value": "Entertainment"
                            },
                            {
                              "title": "Transportation",
                              "value": "Transportation"
                            },
                            {
                               "title": "Lodging",
                               "value": "Lodging"
                            },  
                            {
                               "title": "Misc",
                               "value": "Misc"
                            },                      
  
                          ]
                      },
                      
                      {
                        "type": "TextBlock",
                        "text": "Amount",
                        "wrap": true
                      },
                      {
                        "type": "Input.Number",
                        "id": "expAmount",
                        "placeholder": "Amount",
                       
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
                      'type': 'expFormSubmit'
                  }
              }
          ]


    
  }
}
return expenseFormCard;
}