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
    if (req.user) {
        res.render('home', {
            authenticated: true
        })
    } else {
        res.render('home', {
            authenticated: false
        })
    }
})

app.get('/login',
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
    req.logOut()
    res.redirect('/')
})

app.listen(port);
console.log('Server started')