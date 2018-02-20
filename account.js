const assert = require('assert');
const discover = require('./discovery')
const exchange = require('./exchange')
const adapter = new (require('./redis_adapter'))('Account')
const debug = require('util').debuglog("account")
const aguid = require('aguid')
const crypto = require('crypto')

class Account {
  constructor(user) {
    this.user = user;
    this.accountId = aguid(this.user.email.toLowerCase().trim()).toString()
  }

  claims() {
    return {
      sub: this.accountId,
      address: {
        country:this.user.address.country,
        formatted:`${this.user.address.street}\n${this.user.address.city}, ${this.user.address.state} ${this.user.address.postal}\n${this.user.address.country}`,
        locality:this.user.address.city,
        postal_code:this.user.address.postal,
        region:this.user.address.state,
        street_address:this.user.address.street,
      },
      email:this.user.email.toLowerCase().trim(),
      email_verified: true,
      family_name: this.user.name.last,
      middle_name: '',
      name: this.user.name.full,
      given_name: this.user.name.first,
      nickname: this.user.name.first,
      phone_number: this.user.phone,
      phone_number_verified: false,
      picture: this.user.photo,
      updated_at: Date.now()
    };
  }

  static async findById(ctx, id) {
    return new Account(await adapter.find(id));
  }

  static async authenticate(email, password) {
    assert(password, 'password must be provided');
    assert(email, 'email must be provided');

    email = email.toLowerCase().trim()

    let discovery = await discover(email.substring(email.indexOf('@') + 1))
    if(discovery.exchange) {
      let url = discovery.exchange.host.startsWith('https://') ? discovery.exchange.host : 'https://' + discovery.exchange.host + '/autodiscover/autodiscover.svc'
      if(discovery.exchange.host === 'autodiscover-s.outlook.com') {
        url = null // we already have office365 auto-discovery built into the exchange module.
      }
      debug("account is using discovery url: " + url)
      debug("account is using email: " + email)
      let conn = await exchange(email, password, url)
      let user = await conn.user(email)
      let account = new Account(JSON.parse(JSON.stringify(user)))
      try {
        await adapter.upsert(account.accountId.toString(), user)
      } catch (e) {
        console.error(e)
        console.error(e.previousErrors)
      }
      return account
    } else if (discovery.imaps) {
      throw new Error('Notimplemented')
    } else {
      throw new Error('Unable to find login provider')
    }
  }
}

module.exports = Account;

