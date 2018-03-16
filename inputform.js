
exports.inputform = function () {
    
  
      var inputformCard = {
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
                        "text": "Leave Request Form",
                        "weight": "bolder",
                        "size": "medium"
                      },
                      {
                        "type": "TextBlock",
                        "text": "Leave Request form Description",
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
                        "text": "Subject",
                        "wrap": true
                      },
                      {
                        "type": "Input.Text",
                        "id": "LeaveSubject",
                        "placeholder": "Subject"
                      },
                      {
                        "type": "TextBlock",
                        "text": "Type",
                        "wrap": true
                      },
                      {
                        "type": "Input.ChoiceSet",
                        "id": "LeaveType",
                        "style": "compact",
                        "value": "1",
                        
                        "choices": [
                            {
                              "title": "Vacation",
                              "value": "Vacation"
                            },
                            {
                              "title": "Sick",
                              "value": "Sick"
                            },
                            {
                              "title": "Floating Holiday",
                              "value": "Floating Holiday"
                            },
                            {
                                "title": "Jury Duty",
                                "value": "Jury Duty"
                              },
                              {
                                "title": "Other",
                                "value": "Other"
                              },

                          ]
                      },
                      {
                        "type": "TextBlock",
                        "text": "Enter Start date"
                      },
                      {
                        "type": "Input.Date",
                        "id": "LeaveStartDate",
                       
                        
                      },
                      {
                        "type": "TextBlock",
                        "text": "Enter End date"
                      },
                      {
                        "type": "Input.Date",
                        "id": "LeaveEndDate",
                       
                        
                      },
                      {
                        "type": "TextBlock",
                        "text": "Enter Leave Reason"
                      },
                      {
                        "type": "Input.Text",
                        "id": "LeaveReason",
                        "isMultiline": true,
                        "style": "text",
                        "maxLength": 0,
                        
                      },
                      {
                        "type": "TextBlock",
                        "text": "Enter Manager "
                      },
                      {
                        "type": "Input.Text",
                        "id": "Leavemanager",
                       
                        
                      }
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
                        'type': 'inputFormSubmit'
                    }
                }
            ]


      
    }
  }
  return inputformCard;
  }