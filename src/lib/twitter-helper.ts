const Twitter = require('twitter');
// import * as Twitter from 'twitter';

import { TwitterConsumerKey, TwitterConsumerSecret, TwitterAccessTokenKey, TwitterAccessTokenSecret } from '../config/';

export const twitterClient = new Twitter({
  consumer_key: TwitterConsumerKey,
  consumer_secret: TwitterConsumerSecret,
  access_token_key: TwitterAccessTokenKey,
  access_token_secret: TwitterAccessTokenSecret
});
