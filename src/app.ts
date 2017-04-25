import { twitterClient, Excel } from './lib/';

const excel = new Excel();
const stream = twitterClient.stream('statuses/filter', { track: 'JavaScript' });  // '#MicrosoftGraph'

stream.on('data', ev => {
  excel.updateSheets(ev);
});

stream.on('error', err => {
  throw err;
});
