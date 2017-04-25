
import { WorkbookRange } from '@microsoft/microsoft-graph-types';

import { graphClient, twitterClient } from './lib/';

const driveItemId = '01KLKXRJ26T2MN4VG5VFBY2UIHU5NGG34O';

async function insertSampleData(count: number, name: string, username: string, location: string, tweet: string) {
  const client = await graphClient();

  const sampleData: WorkbookRange = {
    values: [[name, username, location, tweet]]
  };

  return await client
    .api(`/me/drive/items/${driveItemId}/workbook/worksheets/Sheet1/range(address='A${count}:D${count}')`)
    .patch(sampleData, (err, res) => {
      debugger;
    });
}


let count = 0;

const stream = twitterClient.stream('statuses/filter', { track: 'JavaScript' });
// const stream = client.stream('statuses/filter', { track: '#MicrosoftGraph' });

stream.on('data', ev => {
  insertSampleData(
    ++count,
    ev.user.name,
    ev.user.screen_name, 
    ev.user.location, 
    ev.text
  );
});

stream.on('error', err => {
  throw err;
});
