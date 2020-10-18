/**
 * Determines the width of the data range.
 */
const TOTAL_COLUMNS = 8;

/**
 * Determines the timezone used when identifying the date.
 */
const TIMEZONE = 'CST';

/**
 * Represents a cat or a group of cats that someone can foster.
 * The plea email is a list of PleaEntry categorized by headings.
 */
interface PleaEntry {
  animalType: string;
  status: string;
  name: string;
  age: string;
  physicalDescription: string;
  pleaNotes: string;
  photo: string;
  feedingNotes: string;
}

/**
 * Reads every row in the spreadsheet and generates a list of plea entries.
 */
const getPleaEntriesFromSheet = (): PleaEntry[] => {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();

  // 2D array indexed by row then column
  // e.g. dataRange[1, 2] represents row 2, column 3 (C)
  const dataRange = sheet.getRange(1, 1, sheet.getMaxRows(), TOTAL_COLUMNS);
  const dataRangeValues = dataRange.getValues();
  const pleaEntries = dataRangeValues
    .map(
      (row: any[]): PleaEntry => ({
        animalType: row[0],
        status: row[1],
        name: row[2],
        age: row[3],
        physicalDescription: row[4],
        pleaNotes: row[5],
        photo: row[6],
        feedingNotes: row[7],
      })
    )
    .filter((entry) => entry.animalType !== '' && entry.status !== '');

  return pleaEntries;
};

/**
 * Returns today's date in the format MM/dd/yy
 */
const getToday = (): string => {
  const today = new Date();
  return Utilities.formatDate(today, TIMEZONE, 'MM/dd/yy');
};

/**
 * Builds the body of the plea email from a list of plea entries.
 */
const createPleaEmailBody = (pleaEntries: PleaEntry[]): string => {
  const template = HtmlService.createTemplateFromFile('email');
  // Inject `pleaEntries` into the template
  template.pleaEntries = pleaEntries;
  template.heading = 'Syringe Gruelies (3-6 weeks old)';
  const renderedTemplate = template.evaluate().getContent();

  return renderedTemplate;
};

/**
 * Creates a draft foster plea email in the Gmail account associated with this
 * script's spreadsheet.
 */
const createDraftPleaEmail = (
  subjectTitle: string,
  pleaEntries: PleaEntry[]
): void => {
  const to = '';
  const subject = `${subjectTitle} - ${getToday()}`;
  const body = '';
  const options = {
    htmlBody: createPleaEmailBody(pleaEntries),
  };

  const email = GmailApp.createDraft(to, subject, body, options);
};

/**
 * Creates a draft foster plea email for only Syringue Gruelies Neonatal
 * Orphans that need to go on the plea (not on hold).
 */
const createNeoNatalSGDraftPleaEmail = (): void => {
  const filteredPleaEntries = getPleaEntriesFromSheet().filter(
    (entry) =>
      entry.animalType === 'Neonatal Orphan' &&
      entry.status === 'Foster Plea' &&
      entry.feedingNotes === 'SG'
  );

  createDraftPleaEmail('Neonatal Foster Plea', filteredPleaEntries);

  console.log('Done!');
};
