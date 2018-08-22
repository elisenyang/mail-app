' use strict'

const express = require('express');
const session = require('express-session');
const morgan = require('morgan');
const bodyParser = require('body-parser');
const cookieParser = require('cookie-parser');
const path = require('path');
const passport = require('passport');
const OIDCStrategy = require('passport-azure-ad').OIDCStrategy;
const config = require('./utils/config.js');
const MicrosoftGraph = require("@microsoft/microsoft-graph-client");
require('es6-promise').polyfill();
const fetch = require('isomorphic-fetch');
const GoogleStrategy = require('passport-google-oauth').OAuth2Strategy

const app = express();
const port = process.env.PORT || 3000;

//view engine setup
var hbs = require('express-handlebars')({
    extname: '.hbs'
  });
app.engine('hbs', hbs);
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'hbs');

//authentication

const users = [];

var callback = (iss, sub, profile, accessToken, refreshToken, done) => {
    if (!profile.oid) {
      return done(new Error("No oid found"), null);
    }
  
    findByOid(profile.oid, function(err, user){
        if (err) {
          return done(err);
        }
    
        if (!user) {
          users.push({id: 1, outlookProfile: profile, outlookToken: accessToken, refreshToken, 'outlook': true});
          return done(null, profile);
        }

        user.outlookProfile = profile
        user.outlookToken = accessToken
        user.outlook = true
    
        return done(null, user);
    });
  };
  
passport.use(new OIDCStrategy(config.creds, callback))

passport.use(new GoogleStrategy({
    clientID: "955451874654-o9sjnt2mlcmm4qpllahf5oppbc43ttmn.apps.googleusercontent.com",
    clientSecret: "Q9TGF3IKTqVXT2yIwQRazykv",
    callbackURL: "http://localhost:3000/auth/google/callback"
},
    function(accessToken, refreshToken, profile, done) {
        if (!profile.id) {
            return done(new Error("No id found"), null);
          }
        
          findByOid(profile.id, function(err, user){
            if (err) {
              return done(err);
            }
        
            if (!user) {
              users.push({id: 1, googleProfile: profile, googleToken: accessToken, refreshToken, 'google': true});
              return done(null, profile);
            }

            user.googleProfile = profile
            user.googleToken = accessToken
            user.google = true
        
            return done(null, user);
          });
}))

passport.serializeUser((user, done) => { //NEED TO FIX
    if (user.id) {
        done(null, user.id)
    } else {
        done(null, user.oid)
    }
    
  });
  
  passport.deserializeUser((id, done) => {
    findByOid(id, function (err, user) {
        done(err, user);
      });
  });
  
  var findByOid = function(oid, fn) {
    if (users.length > 0) {
        return fn(null, users[0])
    }
    return fn(null, null);
  };


//configurations
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ 'extended': 'false'}));
app.use(cookieParser());
app.use(session({
    secret: 'sshhhhhh',
    name: 'graphNodeCookie',
    resave: false,
    saveUninitialized: false
  }));
app.use(passport.initialize());
app.use(passport.session());

//routes
app.get('/', function(req, res) {
    if (req.isAuthenticated()) {
        //If only signed in with Outlook
        if (req.user.outlook && !req.user.google) {
            fetch('https://graph.microsoft.com/v1.0/me', {
                method: 'GET',
                headers: {
                    "Authorization": 'Bearer ' +req.user.outlookToken
                }
            }).then((res) => {
                return res.json()
            }).then((resp) => {
                res.render('home', {
                    outlookAuth: true,
                    bothAuth: false,
                    googleAuth: false,
                    email: [resp.mail]
                })
            }).catch( err=> {
                console.log(err)
            })
        }

        //If only signed in with Google
        if (req.user.google && !req.user.outlook) {
            fetch('https://www.googleapis.com/gmail/v1/users/'+req.user.googleProfile.id+'/profile', {
                method: 'GET',
                headers: {
                    "Authorization": 'Bearer ' +req.user.googleToken
                }
            }).then((res) => {
                return res.json()
            }).then((resp) => {
                res.render('home', {
                    outlookAuth: false,
                    bothAuth: false,
                    googleAuth: true,
                    email: [resp.emailAddress]
                })
            }).catch( err=> {
                console.log(err)
            })
        }

        //If signed in with both
        if (req.user.google && req.user.outlook) {
            fetch('https://graph.microsoft.com/v1.0/me', {
                method: 'GET',
                headers: {
                    "Authorization": 'Bearer ' +req.user.outlookToken
                }
            }).then((res) => {
                return res.json()
            }).then((resp1) => {
                fetch('https://www.googleapis.com/gmail/v1/users/'+req.user.googleProfile.id+'/profile', {
                    method: 'GET',
                    headers: {
                        "Authorization": 'Bearer ' +req.user.googleToken
                    }
                }).then((res) => {
                    return res.json()
                }).then((resp2) => {
                    res.render('home', {
                        outlookAuth: false,
                        bothAuth: true,
                        googleAuth: false,
                        email: [resp1.mail, resp2.emailAddress]
                    })
                }).catch( err=> {
                    console.log(err)
                })
            }).catch( err=> {
                console.log(err)
            })
        }
    } else {
        res.render('home', {
            authenticated: false
        })
    }
})


app.get('/mail', function(req,res) {
    //If only signed in with Outlook
    if (req.user.outlook && !req.user.google) {
        fetch('https://graph.microsoft.com/v1.0/me/messages?$filter=isRead eq false', {
            method: 'GET',
            headers: {
                "Authorization": 'Bearer ' +req.user.outlookToken
            }
        }).then((res) => {
            return res.json()
        }).then((resp) => {
            resp.value.forEach((mssg)=> {
                mssg.outlook = true
                mssg.formattedDate = new Date(mssg.receivedDateTime)
            })
            var allMessages = resp.value.sort(function(a,b) {
                return new Date(b.date) - new Date(a.date)
            })
            res.render('mail', {
                authenticated: true,
                messages: allMessages
            })
        }).catch( err=> {
            console.log(err)
        })
    }

    //only Google
    if (req.user.google && !req.user.outlook) {
        fetch('https://www.googleapis.com/gmail/v1/users/'+req.user.googleProfile.id+'/messages?q=is:unread', {
            method: 'GET',
            headers: {
                "Authorization": 'Bearer ' +req.user.googleToken
            }
        }).then((res) => {
            return res.json()
        }).then((resp) => {


            var fullMssgs = []
            const promises = resp.messages.map((message) => {
               return fetch('https://www.googleapis.com/gmail/v1/users/'+req.user.googleProfile.id+'/messages/'+message.id+'?format=full', {
                    method: 'GET',
                    headers: {
                        "Authorization": 'Bearer ' +req.user.googleToken
                    }
                }).then((res)=> {
                    return res.json()
                }).then((resp) => {
                    resp.gmail = true
                    for (var i=0; i < resp.payload.headers.length; i++) {
                       if (resp.payload.headers[i].name === 'From') {
                           resp.from = resp.payload.headers[i].value
                       }
                       if (resp.payload.headers[i].name === 'Date') {
                        resp.date = resp.payload.headers[i].value
                        }
                        if (resp.payload.headers[i].name === 'Subject') {
                            resp.subject = resp.payload.headers[i].value
                        }
                    }
                    resp.threadId = message.threadId
                    resp.messageId = message.id
                    fullMssgs.push(resp)
                    return;
                })
            })

            Promise.all(promises).then(()=> {
                var allMessages = fullMssgs.sort(function(a,b) {
                    return new Date(b.date) - new Date(a.date)
                })
                res.render('mail', {
                    authenticated: true,
                    messages: allMessages
                })
            })
        }).catch( err=> {
            console.log(err)
        })
    }

    //both
    if (req.user.google && req.user.outlook) {
        const promiseO =  fetch('https://graph.microsoft.com/v1.0/me/messages?$filter=isRead eq false', {
            method: 'GET',
            headers: {
                "Authorization": 'Bearer ' +req.user.outlookToken
            }
        }).then((res) => {
            return res.json()
        }).then((resp) => {
            resp.value.forEach((mssg)=> {
                mssg.outlook = true
                mssg.date = new Date(mssg.receivedDateTime)
            })
            return resp;
        }).catch( err=> {
            console.log(err)
        })

        const promiseG = fetch('https://www.googleapis.com/gmail/v1/users/'+req.user.googleProfile.id+'/messages?q=is:unread', {
            method: 'GET',
            headers: {
                "Authorization": 'Bearer ' +req.user.googleToken
            }
        }).then((res) => {
            return res.json()
        }).then((resp) => {
            return resp
        }).catch( err=> {
            console.log(err)
        })

        Promise.all([promiseO, promiseG]).then((response) => {
            var fullMssgs = []
            const promises = response[1].messages.map((message) => {
               return fetch('https://www.googleapis.com/gmail/v1/users/'+req.user.googleProfile.id+'/messages/'+message.id+'?format=full', {
                    method: 'GET',
                    headers: {
                        "Authorization": 'Bearer ' +req.user.googleToken
                    }
                }).then((res)=> {
                    return res.json()
                }).then((resp) => {
                    resp.gmail = true
                    for (var i=0; i < resp.payload.headers.length; i++) {
                       if (resp.payload.headers[i].name === 'From') {
                           resp.from = resp.payload.headers[i].value
                       }
                       if (resp.payload.headers[i].name === 'Date') {
                        resp.date = new Date(resp.payload.headers[i].value)
                        }
                        if (resp.payload.headers[i].name === 'Subject') {
                            resp.subject = resp.payload.headers[i].value
                        }
                    }
                    fullMssgs.push(resp)
                    return;
                })
            })

            Promise.all(promises).then(()=> {
                const allMessages = [...response[0].value, ...fullMssgs]
                allMessages.sort(function(a,b) {
                    return new Date(b.date) - new Date(a.date)
                })
                res.render('mail', {
                    authenticated: true,
                    messages: allMessages
                })
            })

        })
    }

    if (!req.user.google && !req.user.outlook) {
        res.redirect('/')
    }
})


app.get('/outlookDelete', function(req, res) {
    var client = MicrosoftGraph.Client.init({
        defaultVersion: 'v1.0',
        debugLogging: true,
        authProvider: function(authDone) {
            authDone(null, req.user.outlookToken)
        }
    })

    client.api('/me/messages/'+req.query.id)
        .delete((err, resp) =>{
            if (err) {
                console.log(err)
                return;
            }
            res.redirect('/mail')
        })   
})

app.get('/gmailDelete', function(req,res) {
    fetch('https://www.googleapis.com/gmail/v1/users/'+req.user.googleProfile.id+'/messages/'+req.query.id+'/trash', {
        method: 'POST',
        headers: {
            "Content-Type": "application/json",
            "Authorization": 'Bearer ' +req.user.googleToken
        }
    }).then((res) => {
        return res.json()
    }).then((resp) => {
        res.redirect('/mail')
    }).catch( err=> {
        console.log(err)
    })
})

app.get('/outlookRead', function(req,res) {
    fetch('https://graph.microsoft.com/v1.0/me/messages/'+req.query.id, {
        method: 'PATCH',
        headers: {
            "Content-Type": "application/json",
            "Authorization": 'Bearer ' +req.user.outlookToken
        },
        body: JSON.stringify({"isRead": "true"})
    }).then((res) => {
        return res.json()
    }).then((resp) => {
        res.redirect('/mail')
    }).catch( err=> {
        console.log(err)
    })
})

app.get('/gmailRead', function(req,res) {
    fetch('https://www.googleapis.com/gmail/v1/users/'+req.user.googleProfile.id+'/messages/'+req.query.id+'/modify', {
        method: 'POST',
        headers: {
            "Content-Type": "application/json",
            "Authorization": 'Bearer ' +req.user.googleToken
        },
        body: JSON.stringify({
            "removeLabelIds": [
              "UNREAD"
            ]
          })
    }).then((res) => {
        return res.json()
    }).then((resp) => {
        res.redirect('/mail')
    }).catch( err=> {
        console.log(err)
    })
})

app.get('/outlookFlag', function(req,res) {
    fetch('https://graph.microsoft.com/v1.0/me/messages/'+req.query.id, {
        method: 'PATCH',
        headers: {
            "Content-Type": "application/json",
            "Authorization": 'Bearer ' +req.user.outlookToken
        },
        body: JSON.stringify({
            "isRead": true,
            "flag": {"flagStatus": "flagged"}
        })
    }).then((res) => {
        return res.json()
    }).then((resp) => {
        res.redirect('/mail')
    }).catch( err=> {
        console.log(err)
    })
})


app.get('/gmailFlag', function(req,res) {
    fetch('https://www.googleapis.com/gmail/v1/users/'+req.user.googleProfile.id+'/messages/'+req.query.id+'/modify', {
        method: 'POST',
        headers: {
            "Content-Type": "application/json",
            "Authorization": 'Bearer ' +req.user.googleToken
        },
        body: JSON.stringify({
            "addLabelIds": [
              "STARRED"
            ],
            "removeLabelIds": [
                "UNREAD"
            ]
          })
    }).then((res) => {
        return res.json()
    }).then((resp) => {
        res.redirect('/mail')
    }).catch( err=> {
        console.log(err)
    })
})

app.post('/outlookReply', function(req,res) {
    fetch('https://graph.microsoft.com/v1.0/me/messages/'+req.query.id+'/reply', {
        method: 'POST',
        headers: {
            "Content-Type": "application/json",
            "Authorization": 'Bearer ' +req.user.outlookToken
        },
        body: JSON.stringify({"comment": req.body.replyMessage})
    }).then((resp) => {
        res.redirect('/mail')
    }).catch( err=> {
        console.log(err)
    })
})

app.post('/gmailReply', function(req, res) {
    fetch('https://www.googleapis.com/gmail/v1/users/'+req.user.googleProfile.id+'/messages/'+req.query.id+'?format=full', {
        method: 'GET',
        headers: {
            "Authorization": 'Bearer ' +req.user.googleToken
        }
    }).then((res)=> {
        return res.json()
    }).then((resp) => {
        for (var i=0; i < resp.payload.headers.length; i++) {
            if (resp.payload.headers[i].name === 'From') {
                resp.from = resp.payload.headers[i].value
            }
            if (resp.payload.headers[i].name === 'Date') {
                resp.date = new Date(resp.payload.headers[i].value)
            }
            if (resp.payload.headers[i].name === 'Subject') {
                resp.subject = resp.payload.headers[i].value
            }
            if (resp.payload.headers[i].name === 'To') {
                resp.to = resp.payload.headers[i].value
            }
        }
        return resp
    }).then((resp) => {
        const encodedResponse  = Buffer.from(
            "Content-Type: text/plain; charset=\"UTF-8\"\n" +
            "MIME-Version: 1.0\n" +
            "Content-Transfer-Encoding: 7bit\n" +
            "Subject: "+resp.subject+"\n" +
            "From: "+resp.to+"\n" +
            "To: "+resp.from+"\n" +
            "In-Reply-To: "+resp.from+ "\n"+
            "References: "+resp.id+"\n\n" +
          
            req.body.replyMessage
          , 'binary').toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')

          var prom1 = fetch('https://www.googleapis.com/gmail/v1/users/'+req.user.googleProfile.id+'/messages/send', {
            method: 'POST',
            headers: {
                "Content-Type": "application/json",
                "Authorization": 'Bearer ' +req.user.googleToken
            }, 
            body: JSON.stringify({
                    raw: encodedResponse,
                    threadId: resp.threadId 
                })

            }).then(()=> {
                return;
            }).catch(err => {
                console.log(err)
            })

            var prom2 = fetch('https://www.googleapis.com/gmail/v1/users/'+req.user.googleProfile.id+'/messages/'+req.query.id+'/modify', {
                method: 'POST',
                headers: {
                    "Content-Type": "application/json",
                        "Authorization": 'Bearer ' +req.user.googleToken
                },
                body: JSON.stringify({
                        "removeLabelIds": [
                          "UNREAD"]
                })
                }).then((res) => {
                    return
                }).catch( err=> {
                    console.log(err)
                })
            
            Promise.all([prom1, prom2]).then(()=> {
                res.redirect('/mail')
            })


    }).catch(err => {
        console.log(err)
    })
})


app.get('/outlookLogin',
    passport.authenticate('azuread-openidconnect', {failureRedirect: '/'})
);

app.post('/token',
    passport.authenticate('azuread-openidconnect', { failureRedirect: '/' }
    ),
  function (req, res) {
    res.redirect('/');
  });


app.get('/auth/google',
  passport.authenticate('google', { scope: ["https://www.googleapis.com/auth/plus.me", "https://mail.google.com/"] }));

app.get('/auth/google/callback', 
passport.authenticate('google', { failureRedirect: '/' }),
  function(req, res) {
    res.redirect('/');
});

// app.get('/logout', function(req,res){
//     users.splice(
//         users.findIndex((obj => obj.profile.oid == req.user.profile.oid)), 1);
//       req.session.destroy( () => {
//           req.logOut();
//           res.clearCookie('graphNodeCookie');
//           res.status(200);
//           res.redirect('/');
//       });
// })

app.listen(port);
console.log('Server started on localhost:3000')