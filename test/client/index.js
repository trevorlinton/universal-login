const Strategy = require('openid-client').Strategy;
const express = require('express')
const session = require('express-session')
const passport = require('passport')
const bodyParser = require('body-parser')
const app = express()

let accounts = {}

app.use(session({
  secret: 'nope',
  resave: true,
  saveUninitialized: true,
  name: 'oidclient',
  cookie: {
  	httpOnly: false,
  	secure: false,
  }
}))

app.use(bodyParser.urlencoded({ extended: false }));
app.use(passport.initialize());
app.use(passport.session());

passport.serializeUser(function(user, cb) {
  cb(null, user.sub);
});

passport.deserializeUser(function(id, cb) {
  if(accounts[id]) {
  	cb(null, accounts[id])
  } else {
  	cb('not found', null)
  }
});

const Issuer = require('openid-client').Issuer;

Issuer.discover('http://localhost:9000') // => Promise
  .then(function (issuer) {
    console.log('Discovered issuer', issuer);
  
	const client = new issuer.Client({
	  client_id: 'foo',
	  client_secret: 'bar',
	  redirect_uris:['http://localhost:9002/auth/cb']
	}); // => Client

	const params = {
	  scope:'openid email address profile'
	  // ... any authorization request parameters go here
	  // client_id defaults to client.client_id
	  // redirect_uri defaults to client.redirect_uris[0]
	  // response type defaults to client.response_types[0], then 'code'
	  // scope defaults to 'openid'
	}

	passport.use('oidc', new Strategy({ client, params, passReqToCallback:false, usePKCE:false }, (tokenset, userinfo, done) => {
	  console.log('tokenset', tokenset);
	  console.log('access_token', tokenset.access_token);
	  console.log('id_token', tokenset.id_token);
	  console.log('claims', tokenset.claims);
	  console.log('userinfo', userinfo);
	  accounts[userinfo.sub] = userinfo
	  return done(null, userinfo)
	  /*User.findOne({ id: tokenset.claims.sub }, function (err, user) {
	    if (err) return done(err);
	    return done(null, user);
	  });*/
	}));

	// start authentication request
	// options [optional], extra authentication parameters
	app.get('/auth', passport.authenticate('oidc', [{}]));

	// authentication callback
	app.get('/auth/cb', passport.authenticate('oidc', { successRedirect: '/user', failureRedirect: '/login' }));

	app.get('/user', (req, res) => {
		res.send(req.user)
	})

	app.listen(9002)

});