
exports.StatusForm = function (statusList,status) {
    
  
    var StatusFormCard = {
        'contentType': 'application/vnd.microsoft.card.adaptive',
        'content': {
          "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
          "type": "AdaptiveCard",
          "version": "1.0",
          "body": [
            {
                "type": "TextBlock",
                "text": statusList,
               
                "separator": false,
               
              },
              {
                "type": "TextBlock",
                "width": "auto",
                "size": "extraLarge",
                "color": "accent",
                "text": status,
                "spacing": "none",
                "separator": true
              }
          ]
         


    
  }
}
return StatusFormCard;
}