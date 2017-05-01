import { twitterClient, Excel } from './lib/';

async function main() {
  const excel = new Excel();
  await excel.init();

  const stream = twitterClient.stream('statuses/filter', { track: 'JavaScript' });  // '#MicrosoftGraph'
  stream.on('data', ev => {
    excel.updateSheets(ev);
  });

  stream.on('error', err => {
    throw err;
  });
}

main();
