import { Client } from '@microsoft/microsoft-graph-client';
import { WorkbookRange } from '@microsoft/microsoft-graph-types';

import { getGraphClient, getIndexFrom2dArray, extractLanguage } from './';

export class Excel {
  private driveItemId = '01KLKXRJ26T2MN4VG5VFBY2UIHU5NGG34O';
  private tweetsSheet = 'Sheet1';
  private statisticsSheet = 'Sheet2';

  private graphClient: Client;
  private tweetCount: number = 0;
  private statistics: any[][];

  constructor() {
    this.init();
  }

  private async init() {
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
        return res.text && res.text.length && res.text[0][0] ? res.text.length : 0
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
    if (!matchedLanguages || !matchedLanguages.length) return;

    matchedLanguages.forEach(language => {
      const idx = getIndexFrom2dArray(this.statistics, language);

      let languageCount = Number(this.statistics[idx[0]][idx[1] + 1]);
      this.statistics[idx[0]][idx[1] + 1] = String(languageCount + 1);
    });

    const newStatistics: WorkbookRange = { values: this.statistics };

    return await this.graphClient
      .api(`/me/drive/items/${this.driveItemId}/workbook/worksheets/Sheet2/usedRange`)
      .patch(newStatistics, (err, res) => {
        debugger;
      });
  }

  public updateSheets(ev: any) {
    this.tweetCount++;
    
    const url = `https://twitter.com/${ev.user.screen_name}/status/${ev.id_str}`;

    this.addTweet(
      ev.user.name,                 // name
      ev.user.screen_name,          // username
      ev.user.location,             // location
      url,                          // url
      ev.text                       // tweet
    );

    const matchedLanguages = extractLanguage(ev.text);
    this.updateStatistics(matchedLanguages);
  }
}
