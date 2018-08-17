' use strict'

const express = require('express');
const session = require('express-session');
const morgan = require('morgan');
const bodyParser = require('body-parser');
const cookieParser = require('cookie-parser');
const path = require('path');
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

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ 'extended': 'false'}));
app.use(cookieParser());

//routes
app.get('/', function(req, res) {
    res.render('home')
})

app.get('/token', function(req, res) {
    res.render('token')
})

app.listen(port);
console.log('Server started')