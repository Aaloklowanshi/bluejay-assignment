const XLSX = require('xlsx');
const fs = require('fs');

// Step 1: Take the file as an input
const file_path = './Assignment_Timecard.xlsx'; 

const workbook = XLSX.readFile(file_path);
const sheet_name = workbook.SheetNames[0];
const df = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name]);

//  Converted the given 'Time In' and 'Time Out' strings to Date objects
df.forEach(entry => {
  entry['Time In'] = entry['Time In'] ? new Date(entry['Time In']) : null;
  entry['Time Out'] = entry['Time Out'] ? new Date(entry['Time Out']) : null;
});

//Sort the DataFrame by 'Employee Name' and 'Time In'
df.sort((a, b) => a['Employee Name'].localeCompare(b['Employee Name']) || a['Time In'] - b['Time In']);

//the main logical ppart
let currentEmployee = null;
let consecutiveDaysCount = 0;

for (let i = 0; i < df.length; i++) {
  const entry = df[i];

  if (currentEmployee !== entry['Employee Name']) {
    currentEmployee = entry['Employee Name'];
    consecutiveDaysCount = 1;
  } else {
    const timeInDiff = (entry['Time In'] - df[i - 1]['Time In']) / (1000 * 60 * 60 * 24);
    if (timeInDiff === 1) {
      consecutiveDaysCount++;
    } else {
      consecutiveDaysCount = 1;
    }
  }

  // a) Employees who have worked for 7 consecutive days
  if (consecutiveDaysCount >= 7) {
    console.log(`\nEmployee: ${currentEmployee}`);
    console.log(`Position ID: ${entry['Position ID']}, Time In: ${entry['Time In']}`);
  }

  // b) Employees with less than 10 hours between shifts but greater than 1 hour
  const timeInDiffHours = i > 0 ? (entry['Time In'] - df[i - 1]['Time Out']) / (1000 * 60 * 60) : 0;
  if (timeInDiffHours > 1 && timeInDiffHours < 10) {
    console.log(`\nEmployee: ${currentEmployee}`);
    console.log(`Position ID: ${entry['Position ID']}, Time In: ${entry['Time In']}`);
  }

  // c) Employees who have worked for more than 14 hours in a single shift
  const timecardHours = entry['Timecard Hours (as Time)'] ? parseFloat(entry['Timecard Hours (as Time)'].split(':')[0]) : 0;
  if (timecardHours > 14) {
    console.log(`\nEmployee: ${currentEmployee}`);
    console.log(`Position ID: ${entry['Position ID']}, Time In: ${entry['Time In']}`);
  }
}
