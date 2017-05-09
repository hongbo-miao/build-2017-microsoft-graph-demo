import { Client } from '@microsoft/microsoft-graph-client';
import { WorkbookRange } from '@microsoft/microsoft-graph-types';

import { DriveItemId } from '../config/';
import { getGraphClient, extractLanguage } from './';

export class Excel {
  private tweetTable = 'Table1';
  private statisticTable = 'Table2';

  private graphClient: Client;
  private statistics = { };

  public async init() {
    this.graphClient = await getGraphClient();

    await this.initStatistics();
  }

  private async addTweet(name: string, username: string, location: string, url: string, tweet: string) {
    const tweetItem = {
      index: null,
      values: [[name, username, location, url, tweet]]
    };

    return await this.graphClient
      .api(`/me/drive/items/${DriveItemId}/workbook/tables/${this.tweetTable}/rows/add`)
      .version('beta')
      .post(tweetItem, (err, res) => {
        debugger;
      });
  }

  private async initStatistics() {
    return await this.graphClient
      .api(`/me/drive/items/${DriveItemId}/workbook/tables/${this.statisticTable}/rows`)
      .version('beta')
      .get()
      .then(res => {
        if (!res || !res.value) return;

        res.value.forEach(statistic => {
          const [language, count] = statistic.values[0];
          this.statistics[language] = count;
        });

        return;
      });
  }

  private async updateStatistics(matchedLanguages: string[]) {
    if (!matchedLanguages || !matchedLanguages.length) return;

    matchedLanguages.map(async language => {
      this.statistics[language]++;
    });

    // convert object to array
    // { 'a': 1, 'b': 2 } -> [['a', 1], ['b', 2]]
    const arrayStatistics = Object.keys(this.statistics).map(e => {
      return [e, this.statistics[e]];
    });

    const newStatistics: WorkbookRange = { values: arrayStatistics };

    return await this.graphClient
      .api(`/me/drive/items/${DriveItemId}/workbook/tables/${this.statisticTable}/databodyrange`)
      .patch(newStatistics, (err, res) => {
        debugger;
      });
  }

  public async updateSheets(ev: any) {
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
}
