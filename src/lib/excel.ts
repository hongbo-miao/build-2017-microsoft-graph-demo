import { Client } from '@microsoft/microsoft-graph-client';
import { WorkbookRange, Message } from '@microsoft/microsoft-graph-types';

import { getGraphClient, getIndexFrom2dArray, extractLanguage } from './';

export class Excel {
  private driveItemId = '01KLKXRJ26T2MN4VG5VFBY2UIHU5NGG34O';
  private tweetsSheet = 'Sheet1';
  private statisticsSheet = 'Sheet2';
  private chart = 'Chart 1';
  // private chartBase64 = '';
  // private email = 'example@mail.com';

  private graphClient: Client;
  private tweetCount: number = 0;
  private statistics: any[][];

  public async init() {
    this.graphClient = await getGraphClient();

    this.tweetCount = await this.getTweetCount();
    this.statistics = await this.getStatistics();
  }

  private async getTweetCount() {
    return await this.graphClient
      .api(`/me/drive/items/${this.driveItemId}/workbook/worksheets/${this.tweetsSheet}/usedRange`)
      .version('beta')
      .get()
      .then(res => {
        // empty sheet is be [['']]
        return res.text && res.text.length && res.text[0][0] ? res.text.length : 0;
      });
  }

  private async addTweet(name: string, username: string, location: string, url: string, tweet: string) {
    
    const tweetItem: WorkbookRange = {
      values: [[name, username, location, url, tweet]]
    };

    return await this.graphClient
      .api(`/me/drive/items/${this.driveItemId}/workbook/worksheets/${this.tweetsSheet}/range(address='A${this.tweetCount}:E${this.tweetCount}')`)
      .patch(tweetItem, (err, res) => {
        debugger;
      });
  }

  private async getStatistics() {
    return await this.graphClient
      .api(`/me/drive/items/${this.driveItemId}/workbook/worksheets/${this.statisticsSheet}/usedRange`)
      .version('beta')
      .get()
      .then(res => res.text);
  }

  private async updateStatistics(matchedLanguages: string[]) {
    if (!matchedLanguages || !matchedLanguages.length || !this.statistics || !this.statistics.length) return;

    matchedLanguages.forEach(language => {
      const idx = getIndexFrom2dArray(this.statistics, language);

      if (!idx || !idx.length) return;

      let languageCount = Number(this.statistics[idx[0]][idx[1] + 1]);
      this.statistics[idx[0]][idx[1] + 1] = String(languageCount + 1);
    });

    const newStatistics: WorkbookRange = { values: this.statistics };

    return await this.graphClient
      .api(`/me/drive/items/${this.driveItemId}/workbook/worksheets/${this.statisticsSheet}/usedRange`)
      .patch(newStatistics, (err, res) => {
        debugger;
      });
  }

  public async updateSheets(ev: any) {
    this.tweetCount++;
    
    const url = `https://twitter.com/${ev.user.screen_name}/status/${ev.id_str}`;

    await this.addTweet(
      ev.user.name,                 // name
      ev.user.screen_name,          // username
      ev.user.location,             // location
      url,                          // url
      ev.text                       // tweet
    );

    const matchedLanguages = extractLanguage(ev.text);
    await this.updateStatistics(matchedLanguages);
  }

//   public async getChart() {
//     return await this.graphClient
//       .api(`/me/drive/items/${this.driveItemId}/workbook/worksheets/${this.statisticsSheet}/charts/${this.chart}/Image`)
//       .version('beta')
//       .get()
//       .then(res => {
//         console.log('res.value', res.value);
//         return res.value;
//       });
//   }

//   public async sendMail() {
//     console.log('image', this.chartBase64);

//     let message: Message = {
//       subject: 'Microsoft Graph TypeScript Sample',
//       toRecipients: [{
//         emailAddress: {
//           address: this.email
//         }
//       }],
//       body: {
//         content: `
// <h1>Microsoft Graph</h1>
// <p>This is the report</p>

// <img src="data:image/png;base64,${this.chartBase64}">
// `,
//         contentType: 'html'
//       }
//     };

//     return await this.graphClient
//       .api('/users/me/sendMail')
//       .post({ message })
//       .then(res => {
//         console.log('Mail sent!')
//       }).catch(err => {
//         debugger;
//       });
//   }
}
