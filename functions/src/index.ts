/**
 * Import function triggers from their respective submodules:
 *
 * import {onCall} from "firebase-functions/v2/https";
 * import {onDocumentWritten} from "firebase-functions/v2/firestore";
 *
 * See a full list of supported triggers at https://firebase.google.com/docs/functions
 */

import * as admin from 'firebase-admin'
import * as config from '../../config/config.json'
import * as functions from 'firebase-functions'
import * as gmailHelper from '../utils/gmailHelper'
import * as logger from 'firebase-functions/logger'

import { gmail_v1 } from 'googleapis/build/src/apis/gmail'

admin.initializeApp()

interface IMessage {
  id?: string | null | undefined
  threadId?: string | null | undefined
}

interface IEmail {
  id: string | null | undefined
  from: string | null | undefined
  to: string | null | undefined
  subject: string | null | undefined
  snippet: string | null | undefined
  bodyText: string | null | undefined
  bodyHtml: string | null | undefined
}

var GMAIL: gmail_v1.Gmail

exports.setupEmailWatch = functions.auth.user().onCreate(async (user) => {
  const { uid, email } = user
  GMAIL = await gmailHelper.getAuthenticatedGmail()
  const res = await GMAIL.users.watch({
    userId: 'me',
    requestBody: {
      labelIds: ['INBOX', 'UNREAD'],
      topicName: 'projects/realtime-receiving-email/topics/receiving-emails',
    },
  })

  await admin.firestore().collection('emailWatches').doc(uid).set({
    emailAddress: email,
    historyId: res.data.historyId,
  })
})

exports.processEmail = functions.pubsub
  .topic(config.gcp.pubsub.topicName)
  .onPublish(async (message) => {
    GMAIL = await gmailHelper.getAuthenticatedGmail()
    const { emailAddress, historyId } = JSON.parse(
      Buffer.from(message.data, 'base64').toString()
    )
    await admin
      .firestore()
      .collection('emailWatches')
      .doc(emailAddress)
      .update({
        historyId,
      })

    await fetchMessageFromHistory(historyId)
  })

async function fetchMessageFromHistory(historyId: string): Promise<void> {
  const res = await gmailHelper.getHistoryList({
    userId: 'me',
    startHistoryId: historyId,
  })

  const { history } = res.data

  if (history == null || history.length === 0) {
    logger.warn('Does not have any history yet')
    return
  }

  var messages: IMessage[] = []

  history.forEach((item) => {
    const labelsAdded = item.labelsAdded
    const messagesAdded = item.messagesAdded
    if (labelsAdded != null) {
      pushNewLabelAddedMessage(labelsAdded, messages)
    }

    if (messagesAdded != null) {
      pushNewMessageAddedMessage(messagesAdded, messages)
    }
  })

  if (messages.length > 0) {
    messages = messages.reduce((newArr: IMessage[], current: IMessage) => {
      const x = newArr.find(
        (item) => item.id === current.id || item.threadId === current.threadId
      )
      if (!x) {
        return newArr.concat([current])
      }
      return newArr
    }, [])
  }

  for (let index = 0; index < messages.length; index++) {
    const messageId = messages[index]?.id
    const currentMessage = await gmailHelper.getMessageData(
      messageId?.toString()
    )

    if (currentMessage == null || currentMessage == undefined) {
      logger.warn('Message object is null - id: ', messageId)
      continue
    }
    await extractEmailInformation(currentMessage.data, messageId?.toString())
  }
}

function pushNewMessageAddedMessage(
  messagesAdded: gmail_v1.Schema$HistoryMessageAdded[],
  messages: IMessage[]
) {
  for (let index = 0; index < messagesAdded.length; index++) {
    messages.push({
      id: messagesAdded[index].message?.id,
      threadId: messagesAdded[index].message?.threadId,
    })
  }
}

function pushNewLabelAddedMessage(
  labelsAdded: gmail_v1.Schema$HistoryLabelAdded[],
  messages: IMessage[]
) {
  for (let index = 0; index < labelsAdded.length; index++) {
    if (
      labelsAdded[index].labelIds?.some(
        (labelAdd) => ['INBOX', 'UNREAD'].indexOf(labelAdd) >= 0
      )
    ) {
      messages.push({
        id: labelsAdded[index].message?.id,
        threadId: labelsAdded[index].message?.threadId,
      })
    }
  }
}

async function extractEmailInformation(
  message: gmail_v1.Schema$Message,
  messageId: string | undefined
): Promise<void> {
  try {
    const payload = message.payload
    const headers = payload?.headers
    const parts = payload?.parts
    const emailType = payload?.mimeType
    if (headers == null || headers == undefined) {
      logger.warn('Header is not defined')
      return
    }

    var email: IEmail = {
      id: message.id,
      snippet: message.snippet,
      from: '',
      to: '',
      subject: '',
      bodyHtml: '',
      bodyText: '',
    }

    if (emailType?.includes('plain')) {
      email.bodyText = payload?.body?.data
    } else {
      if (parts == null || parts == undefined) {
        logger.debug(
          `Parts is not defined for message id: ${messageId} - mimeType: ${emailType}`
        )
        email.bodyText = payload?.body?.data
      } else {
        parts.forEach((part) => {
          const mimeType = part.mimeType
          switch (mimeType) {
            case 'text/plain':
              email.bodyText = part.body?.data
              break
            case 'text/html':
              email.bodyHtml = part.body?.data
              break
          }
        })
      }
    }

    headers.forEach((header) => {
      const name = header.name
      mapHeaderToEmailHeader(name, email, header)
    })

    await admin.firestore().collection('emails').add(email)
  } catch (err) {
    throw new Error('process email error: ' + err)
  }
}

function mapHeaderToEmailHeader(
  name: string | null | undefined,
  email: IEmail,
  header: gmail_v1.Schema$MessagePartHeader
) {
  switch (name) {
    case 'To':
      email.to = header.value
      break
    case 'From':
      email.from = header.value
      break
    case 'Subject':
      email.subject = header.value
      break
  }
}
