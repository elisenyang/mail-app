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
var callback = (iss, sub, profile, accessToken, refreshToken, done) => {
    if (!profile.oid) {
      return done(new Error("No oid found"), null);
    }
  
    findByOid(profile.oid, function(err, user){
      if (err) {
        return done(err);
      }
  
      if (!user) {
        users.push({profile, accessToken, refreshToken});
        return done(null, profile);
      }
  
      return done(null, user);
    });
  };
  
  passport.use(new OIDCStrategy(config.creds, callback))

const users = [];

passport.serializeUser((user, done) => {
    done(null, user.oid);
  });
  
passport.deserializeUser((id, done) => {
    findByOid(id, function (err, user) {
        done(err, user);
    });
});
  
  var findByOid = function(oid, fn) {
    for (var i = 0, len = users.length; i < len; i++) {
      var user = users[i];
      if (user.profile.oid === oid) {
        return fn(null, user);
      }
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
        var client = MicrosoftGraph.Client.init({
            defaultVersion: 'v1.0',
            debugLogging: true,
            authProvider: function(authDone) {
                authDone(null, req.user.accessToken)
            }
        })

        var store = {}

        client.api('/me')
            .get()
            .then(() => {
                return res.req.user.profile
            }).then((profile)=> {
                res.render('home', {
                    authenticated: true,
                    displayName: profile.displayName,
                    })
                return;
            })
            .catch(err => {
                console.log(err)
            })

    } else {
        res.render('home', {
            authenticated: false
        })
    }
})


app.get('/outlookMail', function(req,res) {
    var client = MicrosoftGraph.Client.init({
        defaultVersion: 'v1.0',
        debugLogging: true,
        authProvider: function(authDone) {
            authDone(null, req.user.accessToken)
        }
    })

    client.api('/me/messages?$filter=isRead eq false')
        .get((err, resp) => {
            //console.log(resp.value)
            res.render('mail', {
                authenticated: true,
                messages: resp.value
            })
        })
})


app.get('/outlookDelete', function(req, res) {
    var client = MicrosoftGraph.Client.init({
        defaultVersion: 'v1.0',
        debugLogging: true,
        authProvider: function(authDone) {
            authDone(null, req.user.accessToken)
        }
    })

    client.api('/me/messages/'+req.query.id)
        .delete((err, resp) =>{
            if (err) {
                console.log(err)
                return;
            }
            res.redirect('/outlookMail')
        })   
})

app.get('/outlookRead', function(req,res) { ///FIX THIS!!
    // var client = MicrosoftGraph.Client.init({
    //     defaultVersion: 'v1.0',
    //     debugLogging: true,
    //     authProvider: function(authDone) {
    //         authDone(null, req.user.accessToken)
    //     }
    // })


    // client.api('/me/messages/'+req.query.id)
    //     .header("content-type", "application/json")
    //     .patch({message: { isRead: true } }, (err,resp) => {
    //         if (err) {
    //             console.log(err)
    //         }
    //         console.log(resp)
    //         res.redirect('/outlookMail')
    //     })

    fetch('https://graph.microsoft.com/v1.0/me/messages'+req.query.id, {
        method: 'PATCH',
        headers: {
            "Content-Type": "application/json",
            "Authorization": 'Bearer ' +req.user.accessToken
        },
        body: JSON.stringify({"isRead": "true"})
    }).then((res) => {
        return res.json()
    }).then((resp) => {
        console.log(resp)
    }).catch( err=> {
        console.log(err)
    })
})

app.get('/outlookFlag', function(req,res) { ///FIX THIS!!
    // var client = MicrosoftGraph.Client.init({
    //     defaultVersion: 'v1.0',
    //     debugLogging: true,
    //     authProvider: function(authDone) {
    //         authDone(null, req.user.accessToken)
    //     }
    // })

    // client.api('/me/messages/'+req.query.id)
    //     .header("content-type", "application/json")
    //     .patch({"flag": {"flagStatus": "flagged"}}, (err,resp) => {
    //         if (err) {
    //             console.log(err)
    //         }
    //         console.log(resp)
    //         res.redirect('/outlookMail')
    //     })

    fetch('https://graph.microsoft.com/v1.0/me/messages'+req.query.id, {
        method: 'PATCH',
        headers: {
            "Content-Type": "application/json",
            "Authorization": 'Bearer ' +req.user.accessToken
        },
        body: JSON.stringify({"flag": {"flagStatus": "flagged"}})
    }).then((res) => {
        return res.json()
    }).then((resp) => {
        res.status(200)
        console.log(resp)
    }).catch( err=> {
        console.log(err)
    })
})

app.get('/outlookReply', function(req,res) {
    fetch('https://graph.microsoft.com/v1.0/me/messages/'+req.query.id, {
        method: 'PATCH',
        headers: {
            "Content-Type": "application/json",
            "Authorization": 'Bearer ' +req.user.accessToken
        },
        body: JSON.stringify({"body" : {
            "content": "TESTING TESTING"
        }})
    }).then((resp) => {
        return resp.json()
    }).then((resp) => {
        res.status(200)
        console.log(resp)
    }).catch( err=> {
        console.log(err)
    })
})

app.get('/outlookCreateReply', function(req,res) {
    fetch('https://graph.microsoft.com/v1.0/me/messages/'+req.query.id+'/createReply', {
        method: 'POST',
        headers: {
            "Content-Type": "application/json",
            "Authorization": 'Bearer ' +req.user.accessToken
        }
    }).then((resp) => {
        return resp.json()
    }).then((resp) => {
        res.status(200)
        res.redirect('/outlookReply?id='+resp.id)
        //console.log(resp)
    }).catch( err=> {
        console.log(err)
    })
})




app.get('/outlookLogin',
  function(req, res, next) {
    passport.authenticate('azuread-openidconnect',
    {
      response: res,
      failureRedirect: '/'
    })(req, res, next);
  },
  function (req, res) {
    res.redirect('/');
});

app.post('/token',
  function(req, res, next) {
    passport.authenticate('azuread-openidconnect',
      {
        response: res,
        failureRedirect: '/'
      }
    )(req, res, next);
  },
  function (req, res) {
    res.redirect('/');
  });

app.get('/logout', function(req,res){
    users.splice(
        users.findIndex((obj => obj.profile.oid == req.user.profile.oid)), 1);
      req.session.destroy( (err) => {
        req.logOut();
      res.clearCookie('graphNodeCookie');
      res.status(200);
      res.redirect('/');
      });
})

app.listen(port);
console.log('Server started on localhost:3000')