export function getIndexFrom2dArray(arr: any[][], val: any): [number, number] {
  for (let i = 0; i < arr.length; i++) {
    const idx = arr[i].indexOf(val);

    if (idx > -1) return [i, idx];
  }
}

const languages = [
  'JavaScript',
  'TypeScript',
  'C#',
  'Python',
  'Swift',
  'Objective-C',
  'PHP',
  'Ruby',
  'Java'
];

export function extractLanguage(tweet: string): string[] {
  let matchedLanguages = [];

  languages.forEach(language => {
    // 'i' means case insensitive, '\\b' means word match
    // take 'Java' for example, 'java', 'Java', '#java' will match, 'JavaScript' won't match
    const regExp = new RegExp(`\\b${language}\\b`, 'i');

    if (tweet.search(regExp) !== -1) matchedLanguages.push(language);
  });

  return matchedLanguages;
}
