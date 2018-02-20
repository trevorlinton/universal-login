
const assert = require('assert')
const path = require('path')
const Provider = require('oidc-provider')
const bodyParser = require('body-parser')
const pkg = require('./package.json');
const RedisAdapter = require('./redis_adapter')

function save_email(req, res, next) {
  if(!req.body || !req.body.email || req.body.email.indexOf('@') === -1) {
    return res.redirect('login?error=' + encodeURIComponent("Please enter a valid email address."))
  }
  req.session.email = req.body.email
  req.session.save(() => {
    next()
  })
}

function csrf(req, res, next) {
  req.session.csrf = Math.random() * 10000000;
  req.session.save(() => {
    next()
  })  
}

function check_csrf(req, res, next) {
  if(!req.session || !req.session.csrf || req.session.csrf.toString().toLowerCase() !== req.body.csrf.toString().toLowerCase()) {
    return res.sendStatus(403)
  }
  csrf(req, res, next)
}

let clients = {
  'foo':{
    'name':'Fugazi',
    'website':'https://fugazi.com'
  }
}

let scope_names = {
  'openid':'Full Name',
  'email':'Email Address',
  'address':'Address',
  'profile':'Photo and profile information'
}

module.exports = async function(expressApp, renderError) {
  assert(process.env.URL, 'process.env.URL is missing')
  assert(process.env.SECURE_KEY, 'process.env.SECURE_KEY missing');
  assert.equal(process.env.SECURE_KEY.split(',').length, 2, 'process.env.SECURE_KEY format invalid');
  assert(process.env.REDIS_URL, 'process.env.REDIS_URL missing, run `heroku-redis:hobby-dev`');
  assert(process.env.KEYSTORE, 'process.env.KEYSTORE is missing, run `node ./tools/keygen.js`')

  // simple account model for this application, user list is defined like so
  const Account = require('./account');

  const oidc = new Provider(process.env.URL, {

    findById: Account.findById,
    acrValues:['session','urn:mace:incommon:iap:bronze'],
    discovery: {
      service_documentation: process.env.URL,
      version: pkg.version,
    },
    // let's tell oidc-provider we also support the email scope, which will contain email and
    // email_verified claims
    claims: {
      openid: ['sub'],
      address: ['address'],
      email: ['email', 'email_verified'],
      phone: ['phone_number', 'phone_number_verified'],
      profile: ['family_name', 'given_name', 'locale', 'middle_name', 'name',
      'nickname', 'picture', 'updated_at'],
    },
    interactionUrl(ctx) {
      return `/interaction/${ctx.oidc.uuid}`;
    },
    features: {
      claimsParameter: true,
      clientCredentials: true,
      devInteractions: false,
      discovery: true,
      encryption: true,
      introspection: true,
      registration: true,
      registrationManagement: true,
      request: true,
      requestUri: true,
      revocation: true,
      sessionManagement: true,
    },

    // TODO: DOUBLE CHECK THESE
    ttl: {
      AccessToken: 1 * 60 * 60, // 1 hour in seconds
      AuthorizationCode: 10 * 60, // 10 minutes in seconds
      ClientCredentials: 10 * 60, // 10 minutes in seconds
      IdToken: 365 * 60 * 60, // 1 year in seconds
      RefreshToken: 365 * 24 * 60 * 60, // 1 year in seconds
      RegistrationAccessToken: 5 * 365 * 24 * 60 * 60, // 5 years?..
    },

    renderError
  });

  const keystore = JSON.parse((new Buffer(process.env.KEYSTORE, 'base64')).toString('utf8'))

  // TODO: Dynamic clients below
  oidc.initialize({
    keystore,
    clients: [
      {
        application_type: 'web',
        client_id: 'foo',
        client_secret: 'bar',
        grant_types: ['refresh_token', 'authorization_code'],
        redirect_uris: ['https://test.ulogin.cloud/auth/cb']
      },
    ],
    // configure Provider to use the adapter
    adapter: RedisAdapter,
  }).then(() => {
    oidc.app.proxy = true;
    oidc.app.keys = process.env.SECURE_KEY.split(',');
  }).then(() => {

    const parse = bodyParser.urlencoded({ extended: false });

    // TODO: Re-add (or validate if we need it) CSCR
    expressApp.get('/interaction/:grant', async (req, res) => {
      try {
        details = await oidc.interactionDetails(req)
        details.params.client = clients[details.params.client_id]

        const view = (() => {
          switch (details.interaction.reason) {
            case 'consent_prompt':
            case 'client_not_authorized':
              return 'authorize';
            default:
              return 'login';
          }
        })();
        res.render(view, { details, email:req.session.email, csrf:req.session.csrf, error:null });
      } catch (e) {
        console.error(e)
      }
    });

    expressApp.post('/interaction/:grant/confirm', [parse], (req, res) => {
      oidc.interactionFinished(req, res, {
        login: {
          account: req.session.account.accountId,
          acr: 'urn:mace:incommon:iap:bronze',
          remember: !!req.body.remember,
          ts: Math.floor(Date.now() / 1000),
        },
        consent: {}
      });
    })

    expressApp.post('/interaction/:grant/login', [parse, save_email], (req, res) => {
      res.render('password.ejs', {csrf:req.session.csrf, email:req.session.email, error:req.query.error})
    })

    expressApp.post('/interaction/:grant/password', [parse], async (req, res, next) => {
      try {
        let account = await Account.authenticate(req.body.email, req.body.password)
        req.session.account = account
        req.session.save(async () => {
          details = await oidc.interactionDetails(req)
          details.params.client = clients[details.params.client_id]
          res.render('authorize', {details, email:req.session.email, account:req.session.account, scope_names, csrf:req.session.csrf})
        })
      } catch(err) {
        // TODO: Check to see if its a different error, such as unable to discover host. 
        console.log(err)
        res.render('password.ejs', {csrf:req.session.csrf, email:req.session.email, error:'unauthorized'})
      }
    });

    // leave the rest of the requests to be handled by oidc-provider, there's a catch all 404 there
    expressApp.use(oidc.callback);
  })
}