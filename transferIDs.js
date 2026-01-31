/*--------------------------------------------------------------------------------
Script to transfer student IDs from Google Sheet "Destiny Fines", to "ID Database"
Executes checkDay()

** Google Sheet identifiers have been removed and replaced with "IDENTIFIER REMOVED"
--------------------------------------------------------------------------------*/


/*
Main Function
*/
function main() {
	
	// Checks day
	if (isWeekend()) return;

	// Open the Google Sheets used
	// Opens Google Sheet: Destiny Fines
	const destinyFinesID = "IDENTIFIER REMOVED";
	const destinyFinesSheet = SpreadsheetApp.openByID(destinyFinesID);
	// Open Tabs in Destiny Fines
	const noLanyardTab = destinyFinesSheet.getSheetByName("No Lanyard");
	const freeLanyardTab = destinyFinesSheet.getSheetByName("Free Lanyard");
	const mergeTab = destinyFinesSheet.getSheetByName("Merge");
	
	// Opens Google Sheet: ID Database
	const databaseID = "IDENTIFIER REMOVED";
	const databaseSheet = SpreadsheetApp.openByID(databaseID);
	// Open Tabs in ID Database
	const noIDTab = databaseSheet.getSheetByName("ID Reprints - No ID");
	const noLanyardTab = databaseSheet.getSheetByName("ID Reprints - No Lanyard");
	const freeLanyardTab = databaseSheet.getSheetByName("Free Lanyard List");
	
	// Create arrays for each Google Sheet tab
	let mergeTabArray = [];
	let freeLanyardTabArray = [];
	let noLanyardTabArray = [];
	
	// Get todays date to splice with each student id
	const date = getFormattedDate();

	// If Merge Tab is not empty, add values to the database
	if (!tabEmpty(mergeTab, 1)) {
		mergeTabArray = getTabArray(mergeTab, "A", 1); // *Might have #N/A
		
		if (mergeTabArray[0][0] != '#N/A') {
			const noIdRange = noIDTab.getRange('A' + (noIDTab.getLastRow()+1) + ':' + 'B' + (noIDTab.getLastRow()+mergeTabArray.length));
			noIdRange.setValues(spliceDate(mergeTabArray, date));
		}
	}
	
	// If Free Lanyard Tab is not empty, add values to the database
	if (!tabEmpty(freeLanyardTab, 3)) {
		freeLanyardTabArray = getTabArray(freeLanyardTab, "A", 4);
		const freeLanyardRange = freeLanyardTab.getRange('A'+ (freeLanyardTab.getLastRow()+1) + ':' + 'B' + (freeLanyardTab.getLastRow()+freeLanyardTabArray.length));
		freeLanyardRange.setValues(insert_date(free_lanyard_values, date_value));
	}
	
	// If No Lanyard Tab is not empty, add values to the database
	if (!tabEmpty(noLanyardTab, 1)) {
		noLanyardTabArray = getTabArray(noLanyardTab, "A", 2);
		const noLanyardRange = noLanyardTab.getRange('A'+ (noLanyardTab.getLastRow()+1) + ':' + 'B' + (noLanyardTab.getLastRow()+noLanyardTabArray.length));
		noLanyardTab.setValues(spliceDate(noLanyardTabArray, date));
	 }
	
	// Clear contents of Destiny Fines for the next day
	clearTabs(destinyFinesSheet, "A", "A2:A");
	clearTabs(destinyFinesSheet, "B", "A2:A");
	clearTabs(destinyFinesSheet, "ID Reprints - No Lanyard", "A2:A");
	clearTabs(destinyFinesSheet, "Free Lanyard List", "A4:A");
	clearTabs(destinyFinesSheet, "Free Lanyard List", "B2");
	
	// Run Discipline program...
	
	// Send output as email...
	emailNotification();
}

/*
Checks whether the main script should run.
Weekend it will return, Weekday it will run the main script.
*/
function isWeekend() {
	const date = new Date();
	const day = date.getDay();
	(day == 0 || day == 6) ? return true: return false;
}

/*
Checks to see if parameter tab is empty.
Tab is empty if startPos equals the last row of the sheet.
if empty return true, 
else return false
*/
function tabEmpty(tab, startPos) {
	(tab.getLastRow() > startPos) ? return false : return true;
}

/*
Returns an array of values for the given parameters
tab = Tab of a Google Sheet
col = Accepts uppercase alphabet for the correspoding column
row = Accepts integer for the corresponding row
col + row is the starting range of the array
*/
function getTabArray(tab, col, row) {
	const lastRow = tab.getLastRow();
	const range = '${col}${row}:${col}${lastRow}';
	return tab.getRange(range).getValues().flat();
}

/*
Returns date in format of MM/dd/yyy
*/
function getFormattedDate() {
	return Utilities.formatDate(new Date(), "GMT+1", "MM/dd/yyyy");
}

/*
Inserts parameter date into the double array
Returns the double array
*/
function spliceDate(doubleArr, date) {
    for (let i=0; i<doubleArr.length ; ++i) {
      doubleArr[i].splice(0, 0, date);
    }
	return doubleArr;
}

/*
Clears the contents of the given parameters
tab = String
range = String
*/
function clearTabs(sheet, tab, range){
	sheet.getSheetByName(tab).getRange(range).clearContent();
}

/*
The listed email recipients are notified of the script running.
My case was to notify the admins to determine discipline level.
*/
function emailNotification() {
	const recipients = "IDENTIFIER REMOVED";
	const subject = "Updated: Student ID Discipline - Sheet";
	const body = "Hello, this is a notification. The discipline sheet is now up to date.";
	
	MailApp.sendEmail(recipients, subject, body);
}
