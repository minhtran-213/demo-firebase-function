import * as config from '../../config/config.json'

import { gmail_v1, google } from 'googleapis'

import { logger } from 'firebase-functions'

const GoogleKeyPath = config.auth.googleKeyFilePath
const AuthSubject = config.auth.subject
const scopes = config.gmail.scopes
const JWT = google.auth.JWT
const authClient = new JWT({
  keyFile: GoogleKeyPath,
  scopes: scopes,
  subject: AuthSubject,
})
var authenticatedGmail: gmail_v1.Gmail
export async function getAuthenticatedGmail() {
  try {
    if (authenticatedGmail != null) {
      logger.debug('Authenticated!')
      return authenticatedGmail
    }
    await authClient.authorize()

    authenticatedGmail = google.gmail({
      auth: authClient,
      version: 'v1',
    })
    return authenticatedGmail
  } catch (err) {
    logger.warn('Authentication Failed')
    throw err
  }
}

export async function getHistoryList(options: any) {
  try {
    await getAuthenticatedGmail()
    return authenticatedGmail.users.history.list(options)
  } catch (err) {
    logger.warn('Get history list failed: ', err)
    throw err
  }
}

export async function getMessageData(messageId: string | undefined) {
  try {
    await getAuthenticatedGmail()
    return authenticatedGmail.users.messages.get({
      userId: 'me',
      id: messageId,
    })
  } catch (err) {
    logger.warn('Get message detail failed: ', err)
    throw err
  }
}
