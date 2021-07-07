module.exports = async function (context, myTimer) {
    var timeStamp = new Date().toISOString();
    
    if (myTimer.isPastDue)
    {
        context.log('JavaScript is running late!');
    }

    // required for MicrosoftGraph
    require("isomorphic-fetch");

    const MicrosoftGraph = require("../node_modules/@microsoft/microsoft-graph-client/lib/src/index.js");

    const secrets = require("./secrets");

    const fs = require("fs");

    let accessToken = secrets.accessToken;
    let pushoverKey = secrets.pushoverKey;
    let pushoverToken = secrets.pushoverToken;

    // Get a new Access Token from Graph Explorer: https://developer.microsoft.com/en-us/graph/graph-explorer
    const client = MicrosoftGraph.Client.init({
        defaultVersion: "v1.0",
        debugLogging: true,
        authProvider: (done) => {
            done(null, accessToken);
        },
    });

    // Get the name of the authenticated user with promises
    client
        .api("/me")
        .select("displayName")
        .get()
        .then((res) => {
            context.log(res);
        })
        .catch((err) => {
            context.log(err);
        });

    https://graph.microsoft.com/v1.0/me/calendarview?startdatetime=2021-04-22T19:09:20.820Z&enddatetime=2021-04-29T19:09:20.820Z
    // GET all events from now 

    // new Date() = now
    var start = new Date();
    var end = new Date();

    // add 1 day
    end.setDate(end.getDate() + 1);

    // fill in 'start' values
    // pad the time with zeros or the api will throw an error
    var startHours = start.getHours().toString().padStart(2, '0');
    var startMinutes = start.getMinutes().toString().padStart(2, '0');
    var startSeconds = start.getSeconds().toString().padStart(2, '0');

    // don't forget to +1 the month. it's zero-based in JavaScript :)
    var startdatetime = `${start.getFullYear()}-${start.getMonth() + 1}-${start.getDate()}T${startHours}:${startMinutes}:${startSeconds}.0Z`

    // fill in 'end' values
    var enddatetime = `${end.getFullYear()}-${end.getMonth() + 1}-${end.getDate()}T23:59:59.999Z`

    context.log(startdatetime)
    context.log(enddatetime)

    await client
        .api('/me/calendarview?startdatetime='+startdatetime+'&enddatetime='+enddatetime)
        .header('Prefer','outlook.timezone="Europe/Berlin"')
        .get()
        .then((res) => {
            var resultsByStartDate = {};
            for (let i = 0; i < res.value.length; i++) {
                var startDate = res.value[i].start.dateTime;
                var title = res.value[i].subject;
                context.log(`${startDate} => ${title}`);
                resultsByStartDate[startDate] = { "title" : title, "timezone": res.value[i].start.timeZone };
            }
            var keys = Object.keys(resultsByStartDate); // or loop over the object to get the array
            // keys will be in any order
            keys.sort(); // maybe use custom sort, to change direction use .reverse()
            // keys now will be in wanted order
            
            // now we have the next appointment
            var firstTime = keys[0];
            var first = resultsByStartDate[firstTime];
            // use date time formatting here x]
            // firstTime = firstTime.replace('T', ' ').substring(0, firstTime.lastIndexOf('.'));
            
            var firstDate = firstTime.substring(0, firstTime.lastIndexOf('T'));
            var firstTimestamp = firstTime.substring(firstTime.lastIndexOf('T') + 1, firstTime.lastIndexOf('.'));

            context.log("firstDate: " + firstDate)
            context.log("firstTimestamp: " + firstTimestamp)

            var pushoverTitle = "First Appointment (" + firstDate + ")"
            var pushoverMessage = first.title + ": " + firstTimestamp

            context.log("pushoverTitle: " + pushoverTitle)
            context.log("pushoverMessage: " + pushoverMessage)

            // send the http request
            var XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;
            var req = new XMLHttpRequest();
            req.open('POST', `https://api.pushover.net:443/1/messages.json?token=${pushoverToken}&user=${pushoverKey}&title=${encodeURIComponent(pushoverTitle)}&message=${encodeURIComponent(pushoverMessage)}`, true);
            req.setRequestHeader('Content-type', 'application/x-www-form-urlencoded');
            req.send();

            // log the result
            req.onreadystatechange = function() {
                if (req.readyState === 4) {
                    if (req.status === 200) {
                        context.log('success')
                    } else {
                        var response = JSON.parse(req.responseText);
                        if(response.errors) {
                            context.log('Error: ' + response.errors);
                        } else {
                            // Lets blast the user with the response :)
                            context.log('Error: ' + req.responseText);
                        }
                    }
                }
            };
            context.log(`first: ${firstTime} (${first.timezone}) => ${first.title}`);
        })
    .catch((err) => {
        context.log(err);
    });
    context.log('JavaScript timer trigger function ran!', timeStamp);   
};

//"schedule": "0 0 20 * * SUN-THU"