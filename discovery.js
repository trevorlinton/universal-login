const util = require('util')
const resolve_srv_ = util.promisify(require('dns').resolveSrv)
const resolve_ = util.promisify(require('dns').resolveAny)
const resolve_mx_ = util.promisify(require('dns').resolveMx)

function standardize(record) {
  if(record && record.type === 'TXT') return null
  let r = {"host":record.value || record.address || record.name || record.exchange, "port": record.port || 443, ttyl: record.ttyl || 30}
  if(r.host.startsWith("10.")) return null
  return r
}

const resolve_srv = async (domain, type) => { try { return standardize((await resolve_srv_(domain))[0], type) } catch (e) { return null } }
const resolve = async (domain, type) => { try { return standardize((await resolve_(domain))[0], type) } catch (e) { return null } }
const resolve_mx = async (domain, type) => { try { return (await resolve_mx_(domain)) } catch (e) { console.log(e); return null } }


async function discover(domain) {
  // check we have a record for exchange
  let exchange = await resolve_srv(`_autodiscover._tcp.${domain}`) || await resolve(`autodiscover.${domain}`)
  let imaps    = await resolve_srv(`_imaps._tcp.${domain}`) || await resolve(`imaps.${domain}`)
  let ldap   = await resolve_srv(`_ldap._tcp.${domain}`) || await resolve(`ldap.${domain}`)
  let smtp   = await resolve_srv(`_smtp._tcp.${domain}`) || await resolve_srv(`_submission._tcp.${domain}`)

  let mx_dest = await resolve_mx(domain)
  if(mx_dest && mx_dest.filter((x) => x.exchange && 
    (x.exchange.toLowerCase().endsWith("google.com") || x.exchange.toLowerCase().endsWith("googlemail.com"))).length > 0)
  {
    exchange = await resolve_srv(`_autodiscover._tcp.gmail.com`) || await resolve(`autodiscover.gmail.com`)
    imaps    = await resolve_srv(`_imaps._tcp.gmail.com`) || await resolve(`imaps.gmail.com`)
    ldap   = await resolve_srv(`_ldap._tcp.gmail.com`) || await resolve(`ldap.gmail.com`)
    smtp   = await resolve_srv(`_smtp._tcp.gmail.com`) || await resolve_srv(`_submission._tcp.gmail.com`)
  }
  
  if(exchange && exchange.host && exchange.host === 'autodiscover.outlook.com') {
    exchange.host = 'autodiscover-s.outlook.com'
  }
  return {
    domain,
    exchange,
    imaps,
    ldap,
    smtp
  }
}

module.exports = discover;