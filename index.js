const express = require('express')
const body = require('body-parser')
const session = require('express-session')
const helmet = require('helmet')
const ejs = require('ejs')
const expressApp = express()
const debug = require('util').debuglog("index")
const openid = require('./openid.js')
const fs = require('fs')

console.assert(process.env.SECRET)
console.assert(process.env.SECURE_KEY, 'process.env.SECURE_KEY missing, run `heroku addons:create securekey`');
console.assert(process.env.SECURE_KEY.split(',').length === 2, 'process.env.SECURE_KEY format invalid');
console.assert(process.env.REDIS_URL, 'process.env.REDIS_URL missing, run `heroku-redis:hobby-dev`');

expressApp.set('trust proxy', 1)
expressApp.use(session({
  secret: process.env.SECRET,
  resave: false,
  saveUninitialized: true,
    name: 'ulogin',
  cookie: {
  	secure: process.env.TEST_MODE ? false : true,
  	httpOnly: true
  }
}))
expressApp.disable('etag')
expressApp.disable('x-powered-by');
expressApp.set('view engine', 'ejs')
expressApp.use(express.static('public'))
expressApp.use(helmet())

expressApp.get('/', (req, res) => res.render('index.ejs', {}))
expressApp.get('/index.html', (req, res) => res.render('index.ejs', {}))
expressApp.get('/questions.html', (req, res) => res.render('questions.ejs'))
expressApp.get('/register.html', (req, res) => res.render('register.ejs', {csrf:req.session.csrf, user:req.session.user, email:req.session.email}))

let errorTemplate = fs.readFileSync('./views/oops.ejs')
openid(expressApp, (ctx, error) => {
	ctx.type = 'html'
	ctx.body = ejs.render(errorTemplate.toString(), {}, {filename:'./views/oops.ejs'})
}).then(() => {
	expressApp.listen(process.env.PORT || 9000, () => console.log('Listening on port ' + (process.env.PORT || 9000)))
}).catch((e) => { console.error(e) })


