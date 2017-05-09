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
