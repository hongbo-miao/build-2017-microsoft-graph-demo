import * as fetch from 'node-fetch';
import { Headers } from 'node-fetch';
import { Client } from '@microsoft/microsoft-graph-client';

import { MicrosoftAppRefreshToken, MicrosoftAppSecret, MicrosoftAppID, MicrosoftAppUrl } from '../config/'

export async function getGraphClient() {
  const accessToken = await getAccessToken();

  return Client.init({
    authProvider: done => done(null, accessToken)
  });
}

export async function getAccessToken() {
  let body = {
    client_id: MicrosoftAppID,
    scope: 'calendars.readwrite calendars.readwrite.shared contacts.readwrite contacts.readwrite.shared files.readwrite mail.readwrite mail.send mail.send.shared mailboxsettings.readwrite tasks.readwrite user.readbasic.all',
    redirect_uri: MicrosoftAppUrl,
    grant_type: 'refresh_token',
    client_secret: MicrosoftAppSecret,
    refresh_token: MicrosoftAppRefreshToken
  };

  let options: any = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8'
    }
  };

  // search params
  options.body = Object
    .keys(body)
    .map(key => `${encodeURIComponent(key)}=${encodeURIComponent(body[key])}`)
    .join('&');

  return fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', options)
    .then(res => res.json())
    .then(json => json['access_token']);
}
