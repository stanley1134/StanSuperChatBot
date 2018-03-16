
exports.weatherCard = function (imagecode,temp,high,low,city,region,text) {
  var options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
  var today  = new Date();

  


    var wcard = {
        'contentType': 'application/vnd.microsoft.card.adaptive',
        'content': {
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "type": "AdaptiveCard",
    "version": "1.0",
    "speak": "The forecast for Seattle January 20 is mostly clear with a High of 51 degrees and Low of 40 degrees",
    "body": [
      {
        "type": "TextBlock",
        "text": "'"+city+"', '"+region+"'",
        "size": "large",
        "isSubtle": true
      },
      {
        "type": "TextBlock",
        "text": today.toLocaleDateString("en-US",options),
        "spacing": "none"
      },
      {
        "type": "ColumnSet",
        "columns": [
          {
            "type": "Column",
           
            "width": "auto",
            "items": [
              {
                "type": "Image",
                "url": "http://l.yimg.com/a/i/us/we/52/"+imagecode+".gif",
                "size": "small",
                "text":text
              },
              {
                "type": "TextBlock",
                "text":text,
                "size": "small"
                
              }
            ]
          },
         
          {
            "type": "Column",
            "width": "auto",
            "items": [
              {
                "type": "TextBlock",
                "text": temp +" °F",
                "size": "extraLarge",
                "spacing": "none"
              }
            ]
          },
          
          {
            "type": "Column",
            "width": "auto",
            "items": [
              {
                "type": "TextBlock",
                "text": "Hi "+high+"",
                "horizontalAlignment": "left"
              },
              {
                "type": "TextBlock",
                "text": "Lo "+low+"",
                "horizontalAlignment": "left",
                "spacing": "none"
              }
            ]
          }
        ]
      }
    ]
  }
}
return wcard;
}