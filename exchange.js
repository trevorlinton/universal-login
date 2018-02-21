'use strict'

const ews = require('ews-javascript-api');
const util = require('util');
let debug = util.debuglog('exchange');
const envelope = require('envelope')

ews.EwsLogging.DebugLogEnabled = false;

function DictToObj(values, target) {
  if(!values) {
    return {};
  }
  let ret = {};
  values.keys.forEach((key) => {
    ret[key] = values.objects[key].propertyBag ? values.objects[key].propertyBag.items[target] : values.objects[key][target];
  });
  return ret;
}


function create_propertyset(properties, use_id_only) {
  let propertySet = new ews.PropertySet(use_id_only ? ews.BasePropertySet.IdOnly : ews.BasePropertySet.FirstClassProperties, properties);
  return propertySet;
}

let messagePropertySet = create_propertyset([
  ews.EmailMessageSchema.ItemClass,
  ews.EmailMessageSchema.Size,
  ews.EmailMessageSchema.Sender,
  ews.EmailMessageSchema.Subject,
  ews.EmailMessageSchema.ConversationIndex,
  ews.EmailMessageSchema.ConversationTopic,
  ews.EmailMessageSchema.IsRead,
  ews.EmailMessageSchema.From,
  ews.EmailMessageSchema.ToRecipients,
  ews.EmailMessageSchema.CcRecipients,
  ews.EmailMessageSchema.ReceivedBy,
  ews.EmailMessageSchema.DateTimeSent,
  ews.EmailMessageSchema.DateTimeCreated,
  ews.EmailMessageSchema.MimeContent
]);


let threadPropertySet = create_propertyset([
  ews.ConversationSchema.ConversationId,
  ews.ConversationSchema.ConversationTopic,
  ews.ConversationSchema.UniqueRecipients,
  ews.ConversationSchema.UniqueSenders,
  ews.ConversationSchema.UnreadCount,
  ews.ConversationSchema.Categories,
  ews.ConversationSchema.Size,
  ews.ConversationSchema.Importance,
  ews.ConversationSchema.ItemIds,
  ews.ConversationSchema.LastModifiedTime,
  ews.ConversationSchema.Preview
], true);

let userPropertySet = create_propertyset([
  ews.ContactSchema.DisplayName,
  ews.ContactSchema.GivenName,
  ews.ContactSchema.Initials,
  ews.ContactSchema.MiddleName,
  ews.ContactSchema.NickName,
  ews.ContactSchema.CompleteName,
  ews.ContactSchema.CompanyName,
  ews.ContactSchema.EmailAddress,
  ews.ContactSchema.EmailAddresses,
  ews.ContactSchema.PhysicalAddresses,
  ews.ContactSchema.PhoneNumber,
  ews.ContactSchema.PhoneNumbers,
  ews.ContactSchema.AssistantName,
  ews.ContactSchema.Birthday,
  ews.ContactSchema.BusinessHomePage,
  ews.ContactSchema.Children,
  ews.ContactSchema.Companies,
  ews.ContactSchema.ContactSource,
  ews.ContactSchema.Department,
  ews.ContactSchema.Generation,
  ews.ContactSchema.ImAddress,
  ews.ContactSchema.ImAddresses,
  ews.ContactSchema.JobTitle,
  ews.ContactSchema.Manager,
  ews.ContactSchema.Mileage,
  ews.ContactSchema.OfficeLocation,
  ews.ContactSchema.PhysicalAddressCity,
  ews.ContactSchema.PhysicalAddressCountryOrRegion,
  ews.ContactSchema.PhysicalAddressState,
  ews.ContactSchema.PhysicalAddressStreet,
  ews.ContactSchema.PhysicalAddressPostalCode,
  ews.ContactSchema.PostalAddressIndex,
  ews.ContactSchema.Profession,
  ews.ContactSchema.SpouseName,
  ews.ContactSchema.Surname,
  ews.ContactSchema.WeddingAnniversary,
  ews.ContactSchema.HasPicture,
  ews.ContactSchema.PhoneticFullName,
  ews.ContactSchema.PhoneticFirstName,
  ews.ContactSchema.PhoneticLastName,
  ews.ContactSchema.Alias,
  ews.ContactSchema.Notes,
  ews.ContactSchema.Photo,
  ews.ContactSchema.UserSMIMECertificate,
  ews.ContactSchema.MSExchangeCertificate,
  ews.ContactSchema.DirectoryId,
  ews.ContactSchema.ManagerMailbox,
  ews.ContactSchema.DirectReports,
  /* ItemSchema Types, general types */
  ews.ItemSchema.Attachments,
  ews.ItemSchema.ItemId,
  ews.ItemSchema.ItemClass
]);

let meetingPropertySet = create_propertyset([
  /*ews.AppointmentSchema.ICalUid,
  ews.AppointmentSchema.ICalRecurrenceId,
  ews.AppointmentSchema.ICalDateTimeStamp,
  ews.AppointmentSchema.RequiredAttendees,
  ews.AppointmentSchema.OptionalAttendees*/
])

function MSUniqueIDToID(MsID) {
  return new Buffer(MsID, 'base64').toString('hex')
}

function IDToMSUniqueID(Id) {
  return new Buffer(Id, 'hex').toString('base64');
}

function ItemSchemaUrl(item) {
  if(item.ItemClass.indexOf('Meeting') > -1) {
    return '/events/' + MSUniqueIDToID(item.Id.UniqueId);
  } else {
    return '/messages/' + MSUniqueIDToID(item.Id.UniqueId);
  }
}

function UserUrl(name, email) {
  if(!name) {
    return null;
  }
  let proposed_name = name.toLowerCase()
    .replace(/\@[A-z0-9\.]+$/, '')
    .replace(/ /g, '.')
    .replace(/[^\.A-z0-9]/g, '');

  if(!email) {
    return '/users/' + proposed_name;
  } else {
    let email_name = email.substring(0, email.indexOf('@')).toLowerCase();
    if(email_name !== name) {
      return '/users/' + email_name;
    } else {
      return '/users/' + name;
    }
  }
}

function safely(object, prop) {
  try {
    return object[prop];
  } catch (x) {
    return null;
  }
}

function tryProperties(object, properties, def) {
  for(let i=0; i < properties.length; i++) {
    let property = properties[i];
    try {
      if(object[property]) {
        if(Array.isArray(object[property]) && object[property].length > 0) {
          return object[property][0];
        } else if (!Array.isArray(object[property])) {
          return object[property];
        }
      }
    } catch(e) { /* Exchange tends to crash on null traversals */ }
  }
  return def ? def : null;
}

function find_mime_type(envelope, type) {
  if(!envelope) {
    return null;
  }
  if(envelope.header && envelope.header.contentType && envelope.header.contentType.mime.toLowerCase() === type) {
    return envelope['0']; // i know, i know..
  }
  for(let key in envelope) {
    if(key !== 'header' && envelope[key] !== envelope) {
      let result = find_mime_type(envelope[key], type);
      if(result) {
        return result;
      }
    }
  }
  return null;
}

function User(incoming) {
  try {
    let u = {
      "$self":UserUrl(incoming.contact.propertyBag.properties.objects.DisplayName, incoming.mailbox.address),
      "name":{
        "first":incoming.contact.propertyBag.properties.objects.GivenName,
        "last":incoming.contact.propertyBag.properties.objects.Surname,
        "full":incoming.contact.propertyBag.properties.objects.DisplayName
      },
      "company":incoming.contact.propertyBag.properties.objects.CompanyName,
      "phones":incoming.contact.propertyBag.properties.objects.PhoneNumbers ? 
        DictToObj(incoming.contact.propertyBag.properties.objects.PhoneNumbers.entries, 'phoneNumber') : 
        {},
      "addresses":incoming.contact.propertyBag.properties.objects.PhysicalAddresses ? 
        DictToObj(incoming.contact.propertyBag.properties.objects.PhysicalAddresses.entries, 'objects') : 
        {},
      "email":incoming.mailbox.address,
      "title":incoming.contact.propertyBag.properties.objects.JobTitle,
      "department":{
        "name":incoming.contact.propertyBag.properties.objects.Department,
        "location":incoming.contact.propertyBag.properties.objects.OfficeLocation
      },
      "manager":{
        "$ref": (incoming.contact.propertyBag.properties.objects.Manager && typeof(incoming.contact.propertyBag.properties.objects.Manager) === 'string') ? 
                  UserUrl(incoming.contact.propertyBag.properties.objects.Manager, null) : 
                  null,
        "name":{
          "full":incoming.contact.propertyBag.properties.objects.Manager
        },
        "email":incoming.contact.propertyBag.properties.objects.ManagerMailbox ? 
                incoming.contact.propertyBag.properties.objects.ManagerMailbox.address :
                null
      },
      "photo":incoming.contact.propertyBag.properties.objects.Photo ? 
               'data:image/jpg;base64,' + incoming.contact.propertyBag.properties.objects.Photo :
               null,
      "employees":(incoming.contact.propertyBag.properties.objects.DirectReports ? 
          incoming.contact.propertyBag.properties.objects.DirectReports.items.map(x => { return {"$ref":UserUrl(x.name, null)} }) : 
          [])
    }
    if(u.phones && u.phones.MobilePhone) {
      u.phone = u.phones.MobilePhone
    }
    if(u.addresses && u.addresses.Home) {
      u.address = {
        street:u.addresses.Home.Street,
        city:u.addresses.Home.City,
        state:u.addresses.Home.State,
        country:u.addresses.Home.CountryOrRegion,
        postal:u.addresses.Home.PostalCode
      }
    } else if (u.addresses && u.addresses.Business) {
      u.address = {
        street:u.addresses.Business.Street,
        city:u.addresses.Business.City,
        state:u.addresses.Business.State,
        country:u.addresses.Business.CountryOrRegion,
        postal:u.addresses.Business.PostalCode
      }
    }
    return u
  } catch (e) {
    console.error(e.message)
    console.error(e.stack)
    return {
      "$self":"#/error",
      "error":{
        message:e.message,
        stack:e.stack
      }
    }
  }
}

function Contact(incoming) {
  try {
    return {
      "$self":'/contacts/' + MSUniqueIDToID(incoming.propertyBag.properties.objects.Id.UniqueId),
      "$href":incoming.propertyBag.properties.objects.WebClientReadFormQueryString,
      "name":incoming.propertyBag.properties.objects.CompleteName ?
        ({
          "first":incoming.propertyBag.properties.objects.CompleteName.givenName,
          "middle":incoming.propertyBag.properties.objects.CompleteName.middleName,
          "last":incoming.propertyBag.properties.objects.CompleteName.surname,
          "preferred":incoming.propertyBag.properties.objects.CompleteName.nickname
        }) : {},
      "phone":incoming.propertyBag.properties.objects.PhoneNumbers ? 
        DictToObj(incoming.propertyBag.properties.objects.PhoneNumbers.entries, 'phoneNumber') :
        {},
      "size":incoming.propertyBag.properties.objects.Size
    }
  } catch (e) {
    console.error(e.message)
    console.error(e.stack)
    return {
      "$self":"#/error",
      "error":{
        message:e.message,
        stack:e.stack
      }
    }
  }
}

function Thread(conversation, messages) {
  try {
    let thread = {
      "$self":'/threads/' + MSUniqueIDToID(conversation.Id.UniqueId)
    };
    
    try {
      thread["senders"] = conversation.UniqueSenders ? 
                  conversation.UniqueSenders.items.map((x) => { return {"$ref":UserUrl(x, null), "name":{ "full":x }}}) : 
                  [];
    } catch (n) {
      thread["senders"] = [];
    }
    try {
      thread["recipients"] = conversation.UniqueRecipients ? 
                    conversation.UniqueRecipients.items.map((x) => { return {"$ref":UserUrl(x, null), "name":{ "full":x }}}) : 
                    [];
    } catch (n) {
      thread["recipients"] = [];
    }
    thread["importance"] = conversation.Importance ? conversation.Importance.toLowerCase() : 'normal';
    thread["messages"] = messages.map((message) => {
        try {
          return {"$ref":ItemSchemaUrl(message.Items[0])};
        } catch (e2) {
          console.log('error parsing message type:', e2);
          return {"$self":"#/error", "error":{message:e2.message, stack:e2.stack}};
        }
        // If we have a full mime type message body we can do this; otherwise its best to just return the message id.
        //return Message(message.Items[0]);
      });
    thread["subject"] = conversation.Topic || '';
    thread["size"] = conversation.Size;
    thread["read"] = conversation.UnreadCount > 0 ? true : false;
    thread["updated"] = conversation.LastModifiedTime ? new Date(conversation.LastModifiedTime.momentDate).toISOString() : (new Date()).toISOString();
    return thread;
  } catch (e) {
    console.log('error parsing thread:', e);
    return {
      "$self":"#/error",
      "error":{
        message:e ? e.message : 'Unknown',
        stack:e ? e.stack : 'Unknown'
      }
    }
  }
}

function Email(incoming) {
  try {
    let type = 'messages';
    if(incoming.propertyBag.properties.objects.ItemClass.indexOf('Meeting') > -1) {
      type = 'events'
    }

    let sender = tryProperties(incoming, ['Sender', 'From'],  {name:null, address:null, mailboxType:0});
    let receiver = tryProperties(incoming, ['ReceivedBy','ReceivedRepresenting'], {name:null, address:null, mailboxType:0});
    let receiver_onbehalf = tryProperties(incoming, ['ReceivedRepresenting', 'ReceivedBy'],  {name:null, address:null, mailboxType:0});
    let message = {
      "$self":'/' + type + '/' + MSUniqueIDToID(incoming.propertyBag.properties.objects.Id.UniqueId),
      "$href":incoming.propertyBag.properties.objects.WebClientReadFormQueryString,
      "thread":
        incoming.propertyBag.properties.objects.ConversationId ?
        ({
          "$ref":'/threads/' + MSUniqueIDToID(incoming.propertyBag.properties.objects.ConversationId.UniqueId)
        }) : {},
      "sender":{
        "$ref":UserUrl(sender.name, sender.address),
        "name":{
          "full":sender.name
        },
        "email":sender.address,
        "trusted":(sender.mailboxType > 1 ? true : false)
      },
      "recipient":{
        "$ref":UserUrl(receiver.name, sender.address),
        "name":{
          "full":receiver.name
        },
        "email":receiver.address,
        "onbehalf":{
          "name":{
            "full":receiver_onbehalf.name
          },
          "email":receiver_onbehalf.address,
          "trusted":(receiver_onbehalf.mailboxType > 1 ? true : false)
        },
        "trusted":(receiver.mailboxType > 1 ? true : false)
      },
      "sent":new Date(incoming.propertyBag.properties.objects.DateTimeSent.momentDate).toISOString(),
      "created":new Date(incoming.propertyBag.properties.objects.DateTimeCreated.momentDate).toISOString(),
      "importance":incoming.propertyBag.properties.objects.Importance.toLowerCase(), 
      "subject":{
        "original":incoming.ConversationTopic || '',
        "current":incoming.Subject || ''
      },
      "read":safely(incoming,'IsRead'), /* For whatever reason IsRead causes bailouts */
      "size":incoming.Size
    };
    return message;
  } catch (e) {
    console.error('cannot create message:', e);
    return {
      "$self":"#/error",
      "error":{
        message:e.message,
        stack:e.stack
      }
    }
  }
}

function AttachmentContent(parent, parentType, incoming, index, cb) {
  if(!incoming.propertyBag.properties.objects.Attachments || 
      !incoming.propertyBag.properties.objects.Attachments.items[index]) {
    return cb('Not found', null);
  }
  try {
    let attachment = incoming.propertyBag.properties.objects.Attachments.items[index]
    let obj = {
      '$ref':'/' + parentType + '/' + parent + '/attachments/' + index,
      'type':attachment.contentType,
      'name':attachment.name,
      'size':attachment.size,
      'content_id':attachment.contentId,
      'disposition':attachment.isInline ? 'inline' : 'attachment'
    };
    attachment.Load().then((response) => {
      catcher(() => {
        obj.content = response.responses[0].attachment.base64Content;
        cb(null, obj);
      });
    }, (err) => {
      return cb(err, null);
    });
  } catch (e) {
    console.warn('unable to load attachments:', e);
    return cb(e, null);
  }
}

function Attachments(parent, parentType, incoming) {
  if(!incoming.propertyBag.properties.objects.Attachments) {
    return [];
  }
  return incoming.propertyBag.properties.objects.Attachments.items.map((attachment, index) => {
    try {
      let obj = {
        '$ref':'/' + parentType + '/' + parent + '/attachments/' + index,
        'type':attachment.contentType,
        'name':attachment.name,
        'size':attachment.size,
        'content_id':attachment.contentId,
        'disposition':attachment.isInline ? 'inline' : 'attachment'
      };
      return obj;
    } catch (e) {
      console.warn('unable to load attachments:', e);
      return [];
    }
  });
}

function Message(incoming) {
  let msg = Email(incoming);
  msg.message = {
    html:(incoming.propertyBag.properties.objects.Body.text || '')
  }
  if(incoming.propertyBag.properties.objects.MimeContent) {
    try {
      let env = envelope(new Buffer(incoming.propertyBag.properties.objects.MimeContent.content, 'base64'));
      msg.message.text = find_mime_type(env, 'text/plain');
    } catch (e) {
      console.warn('failed to parse mimetype content.', e);
    }
  }
  msg.attachments = Attachments(MSUniqueIDToID(incoming.propertyBag.properties.objects.Id.UniqueId), 'messages', incoming);
  msg.recipients = incoming.propertyBag.properties.objects.ToRecipients ? 
                    incoming.propertyBag.properties.objects.ToRecipients.items.map((x) => { 
                      return {name:x.name, email:x.address}; 
                    }) : {};
  return msg;
}

function Event(incoming, appointment) {
  let msg = Email(incoming);
  msg.message = {
    html:(incoming.propertyBag.properties.objects.Body ? (incoming.propertyBag.properties.objects.Body.text || '') : '')
  }
  if(incoming.propertyBag.properties.objects.MimeContent) {
    try {
      let env = envelope(new Buffer(incoming.propertyBag.properties.objects.MimeContent.content, 'base64'));
      msg.message.text = find_mime_type(env, 'text/plain');
    } catch (e) {
      console.warn('failed to parse mimetype content.', e);
    }
  }
  msg.recipients = incoming.propertyBag.properties.objects.ToRecipients ? 
                    incoming.propertyBag.properties.objects.ToRecipients.items.map((x) => { 
                      return {name:x.name, email:x.address}; 
                    }) : {};
  msg.reminder = {
    show:incoming.propertyBag.properties.objects.ReminderIsSet,
    when:{
      minutes:incoming.propertyBag.properties.objects.ReminderMinutesBeforeStart,
      by:incoming.propertyBag.properties.objects.ReminderDueBy ? 
        new Date(incoming.propertyBag.properties.objects.ReminderDueBy.momentDate).toISOString() : 
        null
    }
  };
  msg.starts = new Date(incoming.Start.momentDate).toISOString();
  msg.ends = new Date(incoming.End.momentDate).toISOString();
  msg.location = incoming.Location;
  msg.recurring = incoming.IsRecurring;
  msg.cancelled = incoming.IsCancelled;
  msg.attendees = {
    requred:incoming.RequiredAttendees,
    optional:incoming.OptionalAttendees
  }
  msg.organizer = incoming.Organizer ?
    {"$ref":'/users/' + incoming.Organizer.name.replace(' ', '.').toLowerCase()} :
    null;

  
  switch(appointment.MyResponseType.trim().toLowerCase()) {
    case 'accept':
      msg.response = 'accepted';
      break;
    case 'decline':
      msg.response = 'declined';
      break;
    case 'tentative':
      msg.response = 'tentative accepted';
      break;
    default:
      msg.response = appointment.MyResponseType.toLowerCase();
      break;
  }

  if(incoming.propertyBag.properties.objects.ItemClass && incoming.propertyBag.properties.objects.ItemClass.indexOf('Resp') > -1) {
    msg.action = incoming.propertyBag.properties.objects.ResponseType.toLowerCase();
    switch(incoming.propertyBag.properties.objects.ResponseType.toLowerCase()) {
      case 'accept':
        msg.action = 'accepted';
        break;
      case 'decline':
        msg.action = 'declined';
        break;
      case 'tentative':
        msg.action = 'tentative accepted';
        break;
    }
  }

  return msg;
}

// this is very simply an annoying property of ews-javascript-api
// if we happen to throw an error its promise structure breaks down
// potentially causing our entire stack to loose its place. This
// catches all errors prints them then does.. nothing really.
function catcher(func) {
  try {
    return func()
  } catch (e) {
    console.error(e.message);
    console.error(e.stack);
  }
}

function parse_auto_discover(response) {
  return catcher(() => {
    debug('-> parse_auto_discover');
    let obj = {};
    for (var _i = 0, _a = response.Responses; _i < _a.length; _i++) {
      var resp = _a[_i];
      for (var setting in resp.Settings) {
        obj[ews.UserSettingName[setting]] = resp.Settings[setting];
      }
    }
    debug('<- parse_auto_discover');
    return obj;
  });
}

function auto_discover(user, pass, auto_discovery_url, callback) {
  auto_discovery_url = auto_discovery_url || "https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc"
  debug("auto_discover got:", auto_discovery_url)
  if(auto_discovery_url === 'https://autodiscover.outlook.com/autodiscover/autodiscover.svc') {
    debug("resetting url, to normalize it")
    auto_discovery_url = 'https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc'
  }
  catcher(() => {
    debug('-> auto_discover');
    var autod = new ews.AutodiscoverService(new ews.Uri(auto_discovery_url));
    try {
      autod.Credentials = new ews.ExchangeCredentials(user, pass);
    } catch (e) {
      return callback(e, null);
    }
    var settings = [
      ews.UserSettingName.InternalEwsUrl,
      ews.UserSettingName.ExternalEwsUrl,
      ews.UserSettingName.UserDisplayName,
      ews.UserSettingName.UserDN,
      ews.UserSettingName.EwsPartnerUrl,
      ews.UserSettingName.DocumentSharingLocations,
      ews.UserSettingName.MailboxDN,
      ews.UserSettingName.ActiveDirectoryServer,
      ews.UserSettingName.CasVersion,
      ews.UserSettingName.ExternalWebClientUrls,
      ews.UserSettingName.ExternalImap4Connections,
      ews.UserSettingName.AlternateMailboxes
    ];
    autod.GetUserSettings([user], settings)
    .then(function (response) {
      if(!response) {
        return callback('No auto-discovery', null);
      }
      callback(null, parse_auto_discover(response));
    }, ((err) => { callback(err, null); }));
    debug('<- auto_discover');
  });
}

function get_connection(url, user, pass) {
  return catcher(() => {
    debug('-> get_connection:' + url);
    var exch = new ews.ExchangeService(ews.ExchangeVersion.Exchange2015);
    exch.Credentials = new ews.ExchangeCredentials(user, pass);
    exch.Url = new ews.Uri(url);
    debug('<- get_connection');
    return exch;
  });
}

let thread_size = 10;

// private
function delete_item(exch, id, callback) {
  catcher(() => {
    debug('-> delete_item');
    exch.DeleteItem(new ews.ItemId(id),
      ews.DeleteMode.MoveToDeletedItems, 
      ews.SendCancellationsMode.SendToNone, 
      ews.AffectedTaskOccurrence.AffectedTaskOccurrence)
      .then((results) => { catcher(() => {
        callback(null, results);
      })}, 
      (err) => { catcher(() => { callback(err,null); }) });
  });
}

// private
function delete_items(exch, ids, callback) {
  catcher(() => {
    debug('-> delete_items');
    exch.DeleteItems(ids.map((id) => { return new ews.ItemId(id); }),
      ews.DeleteMode.MoveToDeletedItems, 
      ews.SendCancellationsMode.SendToNone, 
      ews.AffectedTaskOccurrence.AffectedTaskOccurrence)
      .then((results) => { catcher(() => {
        callback(null, results);
      })}, 
      (err) => { catcher(() => {
        callback(err,null); }) });
  });
}

function create_message_in_thread(exch, message_id, body, cb) {
  catcher(() => {
    debug('-> create_message_in_thread');
    let id = IDToMSUniqueID(message_id);
    ews.EmailMessage.Bind(exch, new ews.ItemId(id), messagePropertySet).then(
      (msg) => { catcher(() => {
        msg.Reply(new ews.MessageBody(body.message), body.recipients === 'all' ? true : false).then(() => {
          cb(null, {"$ref":"/messages/"});
        }, (err) => {
          console.error('unable to process reply:', err);
          cb(err, null); 
        });
        
      }); },
      (err) => { catcher(() => { cb(err, null); }); });
    debug('<- create_message_in_thread');
  });
}

function delete_thread(exch, id, callback) {
  catcher(() => {
    debug('-> delete_thread');
    id = IDToMSUniqueID(id);
    exch.GetConversationItems(new ews.ConversationId(id), ews.PropertySet.FirstClassProperties).then(
      (results) => { catcher(() => { 
        if(!results.ConversationNodes && results.ConversationNodes.items.length !== 1) {
          return callback('No such thread found', null);
        }
        delete_items(exch, results.ConversationNodes.Items.map((x) => { return x.Items[0].Id.UniqueId; }), callback);
      }) }, 
      (err) => { catcher(() => { callback(err,null); }) });
    debug('<- delete_thread');
  });
}

function get_threads(exch, offset, limit, query, callback) {
  catcher(() => {
    debug('-> get_threads');
    let got_conversation_ids = function(conversations) { catcher( () => {
      let got_threads_items = function(threads) { catcher( () => {
        callback(null, threads.responses.map((thread) => { 
          // find conversation
          let conversation = conversations.filter((conversation) => { 
            return conversation.Id.UniqueId === thread.Conversation.ConversationId.UniqueId 
          });
          if(conversation.length !== 1) {
            console.warn('Cannot fetch conversation for thread, ', conversation);
            // TODO: Return error object
          }
          return Thread(conversation[0], thread.Conversation.ConversationNodes.Items); 
        }));
      })};
      let conversationRequests = conversations.map((x) => {
        return new ews.ConversationRequest(x.Id, null); 
      });
      exch.GetConversationItems(conversationRequests, new ews.PropertySet(ews.BasePropertySet.FirstClassProperties), [], ews.ConversationSortOrder.DateOrderAscending)
         .then(got_threads_items, callback);
    })};

    if(query && query.trim() !== '') {
      let view = new ews.ConversationIndexedItemView(limit, offset, ews.OffsetBasePoint.Beginning);
      view.Traversal = ews.ConversationQueryTraversal.Deep;
      exch.FindConversation( 
        view,
        new ews.FolderId(ews.WellKnownFolderName.Root), 
        'subject:"' + query + '"')
          .then(got_conversation_ids, callback);
    } else {
      exch.FindConversation( 
        new ews.ConversationIndexedItemView(limit, offset, ews.OffsetBasePoint.Beginning), 
        new ews.FolderId(ews.WellKnownFolderName.Inbox))
          .then(got_conversation_ids, callback);
    }
    debug('<- get_threads');
  });
}

function get_thread(exch, id, callback) {
  catcher(() => {
    debug('-> get_thread');
    id = IDToMSUniqueID(id);
    exch.GetConversationItems(new ews.ConversationId(id), ews.PropertySet.FirstClassProperties).then(
      (results) => { catcher(() => { 
        if(!results.ConversationNodes && results.ConversationNodes.items.length !== 1) {
          return callback('No such thread found', null);
        }
        callback(null, results);
        //callback(null, results.ConversationNodes.items[0].Items.map((x) => { return Email(x); })); 
      }) }, 
      (err) => { catcher(() => { callback(err,null); }) });
    debug('<- get_thread');
  });
}

function get_attachments(exch, message_id, callback) {
  catcher(() => {
    debug('-> get_attachments');
    let id = IDToMSUniqueID(id);
    ews.EmailMessage.Bind(exch, new ews.ItemId(id), messagePropertySet).then(
      (response) => { catcher(() => { cb(null, Attachments(message_id, 'messages', response)); }); },
      (err) => { catcher(() => { cb(err, null); }); });
    debug('<- get_attachments');
  });
}

function get_attachment(exch, message_id, attachment_index, cb) {
  catcher(() => {
    debug('-> get_attachment');
    let id = IDToMSUniqueID(message_id);
    ews.EmailMessage.Bind(exch, new ews.ItemId(id), messagePropertySet).then(
      (response) => { catcher(() => { AttachmentContent(message_id, 'messages', response, attachment_index, cb); }); },
      (err) => { catcher(() => { cb(err, null); }); });
    debug('<- get_attachment');
  });
}

function get_attachment_content(exch, message_id, attachment_index, cb) {
  catcher(() => {
    debug('-> get_attachment');
    let id = IDToMSUniqueID(message_id);
    ews.EmailMessage.Bind(exch, new ews.ItemId(id), messagePropertySet).then(
      (response) => { catcher(() => { 
        AttachmentContent(message_id, 'messages', response, attachment_index, (err, data) => {
          if(err) {
            cb(err, null, null);
          } else {
            cb(null, data.content ? new Buffer(data.content, 'base64') : new Buffer(0), data.type);
          }
        }); }); },
      (err) => { catcher(() => { cb(err, null); }); });
    debug('<- get_attachment');
  });
}

function get_messages(exch, callback) {
  catcher(() => {
    debug('-> get_messages');
    exch.FindItems(ews.WellKnownFolderName.Inbox, new ews.ItemView(1000)).then(
      (results) => { catcher(() => { callback(null, results.items.map((x) => { return Email(x); })); }) }, 
      (err) => { catcher(() => { callback(err, null); }) });
    debug('<- get_messages');
  });
}

function get_message(exch, id, cb) {
  catcher(() => {
    debug('-> get_message');
    id = IDToMSUniqueID(id);
    ews.EmailMessage.Bind(exch, new ews.ItemId(id), messagePropertySet).then(
      (msg) => { catcher(() => { cb(null, Message(msg)); }); },
      (err) => { catcher(() => { cb(err, null); }); });
    debug('<- get_message');
  });
}

function get_events(exch, callback) {
  catcher(() => {
    debug('-> get_events');
    let current = new Date();
    let week_ahead = new Date();
    week_ahead.setDate(week_ahead.getDate() + 7);
    let view =  new ews.CalendarView(new ews.DateTime(current),new ews.DateTime(week_ahead));
    view.PropertySet = meetingPropertySet;
    exch.FindAppointments(ews.WellKnownFolderName.Calendar, view).then(
      (results) => { catcher(() => { 
        callback(null, results.items.map((x) => {
          return Event(x); 
        })); }); 
      }, 
      (err) => { catcher(() => { 
        callback(err, null); }); 
      });
    debug('<- get_events');
  });
}

function update_event(exch, id, payload, cb) {
  catcher(() => {
    debug('-> update_event');
    id = IDToMSUniqueID(id);
    ews.EmailMessage.Bind(exch, new ews.ItemId(id),  new ews.PropertySet(ews.BasePropertySet.FirstClassProperties)).then(
      (response) => { catcher(() => { 
        ews.Appointment.Bind(exch, response.AssociatedAppointmentId, new ews.PropertySet(ews.BasePropertySet.FirstClassProperties)).then(
          (a_resp) => { catcher(() => {
            switch(payload.response) {
              case 'accept':
              case 'accepted':
                a_resp.Accept(true).then(catcher(() => {
                  let ev = Event(response, {"MyResponseType":"Accept"});
                  ev.response = 'accepted'; 
                  cb(null, ev); 
                }));
                break;
              case 'delined':
              case 'decline':
                a_resp.Decline(true).then(catcher(() => {
                  let ev = Event(response, {"MyResponseType":"Decline"});
                  ev.response = 'declined'; 
                  cb(null, ev); 
                }));
                break;
              case 'tentative accepted':
              case 'tentative':
                a_resp.AcceptTentatively(true).then(catcher(() => {
                  let ev = Event(response, {"MyResponseType":"Tentative"});
                  ev.response = 'tentative accepted'; 
                  cb(null, ev); 
                }));
                break;
              default:
                callback('Response unclear', null);
                break;
            }
          }); },
          (err) => { catcher(() => { cb(err, null); }); })
        
      }); },
      (err) => { catcher(() => { cb(err, null); }); });
    debug('<- update_event');
  });
}

function get_event(exch, id, cb) {
  catcher(() => {
    debug('-> get_event');
    id = IDToMSUniqueID(id);
    ews.MeetingMessage.Bind(exch, new ews.ItemId(id), meetingPropertySet).then(
      (response) => { catcher(() => {
        ews.Appointment.Bind(exch, response.AssociatedAppointmentId, new ews.PropertySet(ews.BasePropertySet.FirstClassProperties)).then(
          (a_resp) => { catcher(() => {
            cb(null, Event(response, a_resp)); 
          }) },
          (err) => { catcher(() => { cb(err, null); }); });
      }); },
      (err) => { catcher(() => { cb(err, null); }); });
    debug('<- get_event');
  });
}

function get_availability(exch, emails, next_hours, callback) {
  catcher(() => {
    next_hours = next_hours || 48;
    var attendee = emails.map((x) => new AttendeeInfo(x));
    var timeWindow = new ews.TimeWindow(ews.DateTime.Now, new ews.DateTime(ews.DateTime.Now.TotalMilliSeconds + ews.TimeSpan.FromHours(next_hours).asMilliseconds())); 
    exch.GetUserAvailability(attendee, timeWindow, ews.AvailabilityData.FreeBusyAndSuggestions)
    .then(callback.bind(callback, null), callback.bind(callback));
  });
}

function search_contacts(exch, name, callback) {
  catcher(() => {
    debug('-> get_user');
    exch.FindItems(ews.WellKnownFolderName.Contacts, 'subject:' + name, new ews.ItemView(1000)).then(
      (results) => { catcher(() => { callback(null, results.items.map((x) => { return Contact(x); })); }) }, 
      (err) => { catcher(() => { callback(err, null); }) });
    debug('<- get_user');
  })
}

function search_users(exch, name, callback) {
  catcher(() => {
    debug('-> search_users');
    if(!name || name === '') {
      callback('No users found.', null);
    }
    exch.ResolveName(name, ews.ResolveNameSearchLocation.DirectoryThenContacts, true, userPropertySet).then(
      (results) => { catcher(() => { callback(null, results.items.map((x) => { return User(x); })); }) },
      (err) => { catcher(() => { callback(err, null); }) });
    debug('<- search_users');
  })
}

function get_user(exch, name, callback) {
  catcher(() => {
    if(!name || name === '') {
      callback('No users found.', null);
    }
    debug('-> get_user');
    exch.ResolveName(name, ews.ResolveNameSearchLocation.DirectoryThenContacts, true, userPropertySet).then(
      (results) => { catcher(() => { 
        if(results.items.length === 0) {
          return callback('not found', null);
        }
        callback(null, User(results.items[0])); 
      })},
      (err) => { catcher(() => { callback(err, null); }); });
    debug('<- get_user');
  });
}

let discovery_cache = {};

function create(user, pass, url, ready) {
  debug('-> create');
  
  function exchange(exch) {
    debug('-> got_connection');
    // public methods
    this.exchange = exch;
    this.thread = get_thread.bind(get_thread, this.exchange);
    this.thread.delete = delete_thread.bind(delete_thread, this.exchange);
    this.threads = get_threads.bind(get_threads, this.exchange);
    this.availability = get_availability.bind(get_availability, this.exchange);
    this.message = get_message.bind(get_message, this.exchange);
    this.messages = get_messages.bind(get_messages, this.exchange);
    this.messages.reply = create_message_in_thread.bind(create_message_in_thread, this.exchange);
    this.attachment = get_attachment.bind(get_attachment, this.exchange);
    this.attachment.content = get_attachment_content.bind(get_attachment_content, this.exchange);
    this.attachments = get_attachments.bind(get_attachments, this.exchange);
    this.event = get_event.bind(get_message, this.exchange);
    this.event.update = update_event.bind(update_event, this.exchange);
    this.events = get_events.bind(get_messages, this.exchange);
    this.search_contacts = search_contacts.bind(search_contacts, this.exchange);
    this.search_users = util.promisify(search_users.bind(search_contacts, this.exchange));
    this.user = util.promisify(get_user.bind(get_user, this.exchange));
    debug('<- got_connection');
  }

  let setup = function() {
    catcher(function() {
      debug('-> setup');
      if(!url && !discovery_cache[user]) {
        auto_discover(user, pass, url, function(err, discovery) {
          catcher(function() {
            if(err || !discovery) {
              return ready(err || 'No discovery was found.');
            }
            discovery_cache[user] = discovery;
            ready(null, new exchange(get_connection(discovery.ExternalEwsUrl, user, pass)));
            debug('<- setup (discovered, no cache)');
          }.bind(this));
        }.bind(this));
      } else if(discovery_cache[user]) {
        ready(null, new exchange(get_connection(discovery_cache[user].ExternalEwsUrl, user, pass)));
        debug('<- setup (cached)');
      } else {
        ready(null, new exchange(get_connection(url, user, pass)));
        debug('<- setup (no cache, no discovery)');
      }
    }.bind(this));
  }.bind(this);
  setup();
  debug('<- create');
}


module.exports = util.promisify(create);
