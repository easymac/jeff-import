import fs from 'node:fs';
import xlsx from 'node-xlsx';

const workSheetsFromFile = xlsx.parse('./Prepared_Sheet.xlsx');

const NSCA_TRAINING_LOADS = {
  1: '100%',
  2: '95%',
  3: '93%',
  4: '90%',
  5: '87%',
  6: '85%',
  7: '83%',
  8: '80%',
  9: '77%',
  10: '75%',
  12: '70%'
}


let result = '';
let currentWeek = 1;
let currentDay = 1;

function parseFile() {
  for (const sheet of workSheetsFromFile) {
    parseSheet(sheet);
  }
}

function parseSheet(sheet) {
  for (const row of sheet.data) {
    if (row.length === 0) continue;
    const stop = parseRow(row);
    if (stop) break;
  }
}

function parseRow(row) {
  // If first cell begins with "Week", start a new week
  if (row[0]?.startsWith('Week')) {
    result += `# Week ${currentWeek}\n`;
    currentWeek++;
    currentDay = 1;
    return;
  }

  // If first cell is 'Mandatory Rest Day', it is definitely not an exercise.
  if (row[0] === 'Mandatory Rest Day') return;

  // If the first cell contains "BLOCK", it is a new block (and gets ignored).
  if (row[0]?.includes('BLOCK')) return;

  // If the first cell contains "DELOAD", it is a note (and gets ignored).
  if (row[0]?.includes('DELOAD')) return;

  // If the row is less than 17 columns, it is definitely not an exercise.
  if (row.length < 17) return;

  parseDay(row);
  parseExercise(row);
}

function parseDay(row) {
  const dayName = row[0];
  if (!dayName) return;
  result += `## Day ${currentDay}\n`;
  currentDay++;
}

function parseExercise(exercise) {
  const exerciseName = exercise[1];
  const url = exercise[2];
  const lastSetIntensity = exercise[3];
  const warmUpSets = exercise[4];
  const workingSets = exercise[6];
  const reps = exercise[7];
  const earlySetRpe = exercise[9];
  const lastSetRpe = exercise[10];
  const rest = exercise[11];
  const notes = exercise[14];

  result += buildDescriptionBlock(exercise);

  result += exerciseName;
  result += ' / ';

  result += buildRepBlock(exercise);
  result += ' / ';

  result += buildUpdateBlock();
  result += ' / ';

  result += buildProgressBlock(exercise);
  result += ' / ';

  result += buildRestBlock(exercise);

  result += '\n\n';
  
}

function buildDescriptionBlock(exercise) {
  let result = '';
  const notes = exercise[17];
  const url = cleanUrl(exercise[2]);
  const lastSetIntensity = exercise[3];

  result += `// ${notes}`;
  if (url) {
    result += '\n//\n';
    result += `// [Exercise video](${url})`;
  }
  if (lastSetIntensity !== 'N/A') {
    result += '\n//\n//\n';
    result += `// **Last set:** ${lastSetIntensity}`;
  }
  result += '\n';
  return result;
}

function buildRepBlock(exercise) {
  const lastSetIntensity = exercise[3];
  let remainingSets = exercise[6];
  const reps = exercise[7];
  const [ min ] = reps.split('-');
  const load = NSCA_TRAINING_LOADS[min];

  let result = '';
  
  // First set
  result += `1x${reps} ${load}+`;
  remainingSets--;

  if (lastSetIntensity !== 'N/A') {
    // Middle sets
    result += `, ${remainingSets - 1}x${reps} ${load}`;

    // Last set
    result += `, 1x${reps} ${load} (Last)`;
  } else {
    // Remaining sets
    result += `, ${remainingSets}x${reps} ${load}`;
  }

  return result;
}

function buildUpdateBlock() {
  return `update: custom() {~ if (setIndex == 1) { weights = weights[1] } ~}`
}

function buildProgressBlock(exercise) {
  const reps = exercise[7];
  const [ min, max ] = reps.split('-');

  if (max) {
    return `progress: dp(5lb, ${min}, ${max})`;
  } else {
    return `progress: lp(5lb)`;
  }
}

function buildRestBlock(exercise) {
  const rest = exercise[14];
  const restWithoutTilde = rest.replace('~', '');
  const [ min, max ] = restWithoutTilde.split('-');

  if (rest === 'N/A') {
    return '0s';
  }

  if (max) {
    return `${min * 60}s`;
  }

  const [ number ] = min.split(' ');
  return `${number * 60}s`;
}

function cleanUrl(url) {
  if (!url) return null;
  // remove the ? and everything after it
  return url.split('?')[0];
}

parseFile();

// Save result as a .txt file
fs.writeFileSync('output.txt', result);