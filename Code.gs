var Sheeturl = "https://docs.google.com/spreadsheets/d/1faZ3pkcrwydPda4zuawftHyDNG4Y8DhUwof1R1BHGzE/edit?gid=0#gid=0";

class Employee {
    constructor(firstName, lastName, id, team, workingHour, skillLevel, actualWorkingHour) {
        this.firstName = firstName;
        this.lastName = lastName;
        this.id = id;
        this.team = team;
        this.workingHour = workingHour;
        this.skillLevel = skillLevel;
        this.actualWorkingHour = actualWorkingHour;
        this.assignedTask = [];
    }
}

class Task {
    constructor(taskName, hoursToWork, priorities) {
        this.taskName = taskName;
        this.hoursToWork = hoursToWork;
        this.priorities = priorities;
    }
}

function setupSheet() {

    var ss = SpreadsheetApp.openByUrl(Sheeturl);
    // Validate if sheet have Employee List and Task List
    var ts = ss.getSheetByName("Task List");
    var es = ss.getSheetByName("Employee List");

    if (ts == null) {
        ts = ss.insertSheet();
        ts.setName("Task List");
        ts.appendRow(["Task", "Department", "Hours to be work on", "Priorities"]);
    }

    if (es == null) {
        es = ss.insertSheet();
        es.setName("Employee List");
        es.appendRow(["First Name", "Last Name", "ID", "Team", "Available hour / week", "Skill level", "Actual hour worked"])
    }
    debugger;
}

/**
 * Get "home page", or a requested page.
 * Expects a 'page' parameter in querystring.
 *
 * @param {event} e Event passed to doGet, with querystring
 * @returns {String/html} Html to be served
 */
function doGet(e) {
    Logger.log(e.parameter);
    if (!e.parameter.page) {
        // When no specific page requested, return to "home page"
        var tmp = HtmlService.createTemplateFromFile('Setup');
        return tmp.evaluate();
    }

    if (e.parameter.page == "Home") {
        var tmp = HtmlService.createTemplateFromFile('Home');
        tmp.sheetUrl = Sheeturl;
        return tmp.evaluate();
    }

    if (e.parameter.page == "Employee List") {
        var ss = SpreadsheetApp.openByUrl(Sheeturl);
        var ws = ss.getSheetByName("Employee List")
        var td = ws.getRange(1, 1, ws.getRange("A1").getDataRegion().getLastRow(), 7).getValues();

        var tmp = HtmlService.createTemplateFromFile('Employee List');
        // tmp.tableData = tableData.map(function(r){ return r[0]; }); 
        tmp.tableData = td;
        return tmp.evaluate();


    }

    if (e.parameter.page == "Task List") {
        var ss = SpreadsheetApp.openByUrl(Sheeturl);
        var ws = ss.getSheetByName("Task List");
        var es = ss.getSheetByName("Employee List");

        if (ws == null) {
            ws = ss.insertSheet();
            ws.setName("Task List");
            ws.appendRow(["Task", "Department", "Hours to be work on", "Priorities"]);
        }

        var tmp = HtmlService.createTemplateFromFile('Task List');
        var ed = es.getRange(1, 1, es.getRange("A1").getDataRegion().getLastRow(), 1).getValues();
        var td = ws.getRange(1, 1, ws.getRange("A1").getDataRegion().getLastRow(), 5).getValues();

        const employeeData = [];
        for (var i = 1; i < employeeData.length; i++) {
            employeeData.push(ed[i][0]);
        }

        tmp.taskData = td;
        tmp.employeeName = employeeData;
        return tmp.evaluate();
    }

    // else, use page parameter to pick an html file from the script
    return HtmlService.createTemplateFromFile(e.parameter['page']).evaluate();
}


/**
 * Get the URL for the Google Apps Script running as a WebApp.
 */
function getScriptUrl() {
    var url = ScriptApp.getService().getUrl();
    return url;
}

function insertNewEmployee(employeeInfo) {
    var ss = SpreadsheetApp.openByUrl(Sheeturl);
    var ws = ss.getSheetByName("Employee List")

    ws.appendRow([employeeInfo.FirstName, employeeInfo.LastName, employeeInfo.ID, employeeInfo.Team, employeeInfo.AvailableHour, employeeInfo.Skill, employeeInfo.ActualHourWorked]);
}

function insertNewTask(task) {
    var ss = SpreadsheetApp.openByUrl(Sheeturl);
    var ws = ss.getSheetByName("Task List")

    ws.appendRow([task.TaskName, task.TaskDepartment, task.WorkHour, task.TaskPriorities]);
}

// Generate resource plan function
function generatePlanSheet() {
    var ss = SpreadsheetApp.openByUrl(Sheeturl);
    var ws = ss.getSheetByName("Resource Plan");
    var ts = ss.getSheetByName("Task List");
    var es = ss.getSheetByName("Employee List");


    var td = ts.getRange(2, 1, ts.getRange("A1").getDataRegion().getLastRow(), 5).getValues();
    var ed = es.getRange(2, 1, es.getRange("A1").getDataRegion().getLastRow(), 7).getValues();


    // Create arrays of employee object 
    var employeeData = [];
    for (var i = 0; i < ed.length; i++) {
        employeeData.push(new Employee(ed[i][0], ed[i][1], ed[i][2], ed[i][3], ed[i][4], ed[i][5]));
    }

    // Create arrays of task object
    var taskData = [];
    for (var i = 0; i < td.length; i++) {
        taskData.push(new Task(td[i][0], td[i][2], td[i][3]));
    }

    // Assgined task to employee
    for (var i = 0; i < td.length; i++) {
        employeeData.forEach(function (arrayItem) {
            if (arrayItem.firstName == td[i][4]) {
                arrayItem.assignedTask.push(td[i][0])
            }
        });
    }

    // Setup Rules


    // Format table
    ws.clear();
    var headers1 = ws.getRange('A1:C1');
    var headers2 = ws.getRange('D1:AI1');
    var table = ws.getDataRange();

    headers1.setFontWeight('bold');
    headers1.setFontColor('white');
    headers1.setBackground('#48489c');

    headers2.setFontWeight('bold');
    headers2.setFontColor('black');
    headers2.setBackground('#e6e3e3');

    var today = new Date();
    for (var i = 0; i < 31; i++) {
        var currentDate = new Date(today);
        currentDate.setDate(today.getDate() + i);
        var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "dd MMM");
        var cell = ws.getRange(1, 4 + i);
        cell.setValue(formattedDate);

    };


    table.setFontFamily('Roboto');
    table.setHorizontalAlignment('center');
    ws.getRange('A1:C1').setValues([['Resource', 'Team', 'Total Effort']]);

    // Print all the working employee
    employeeData.forEach(function (arrayItem) {
        if (arrayItem.assignedTask.length != 0) {
            var headers = ws.getRange(ws.getRange("A1").getDataRegion().getLastRow() + 1, 1, 1, 3); // Get last row for styling 
            headers.setFontWeight('bold');
            headers.setFontColor('black');
            headers.setBackground('#e6e3e3');

            ws.appendRow([arrayItem.firstName, arrayItem.team, arrayItem.workingHour]);
            for (var i = 0; i < arrayItem.assignedTask.length; i++) {
                ws.appendRow([arrayItem.assignedTask[i]]);

            }
        }
    });


    // Hour allocation
    var s = ws.getRange(1, 1, ws.getRange("A1").getDataRegion().getLastRow(), 1).getValues();
    let dateCell = 0;
    let priority = 10;
    let maximumPplPerDay = 2;
    let ad = ts.getRange(19, 7, 3).getValues().filter(String);
    taskData.forEach(function (arrayItem) {

        dateCell = 4;
        priority = 10;
        while (priority >= 0) {
            if (arrayItem.priorities == priority) {
                // Look for the task in the A1 column
                for (var i = 0; i < s.length; i++) {
                    // If task name match we know we want to add hour in that row
                    if (s[i][0] == arrayItem.taskName) {

                        // Continue to allot hour until reach 0 
                        while (arrayItem.hoursToWork != 0) {
                            if (ws.getRange(3, dateCell, ws.getRange("A1").getDataRegion().getLastRow()).getValues().filter(String).length < 2) {
                                if (arrayItem.hoursToWork > 5) {
                                    // remaining hour for task more than 5
                                    ws.getRange(i + 1, dateCell).setValue(5);
                                    arrayItem.hoursToWork = arrayItem.hoursToWork - 5;
                                } else {
                                    //remaining hour for task less than 5
                                    ws.getRange(i + 1, dateCell).setValue(arrayItem.hoursToWork);
                                    arrayItem.hoursToWork = 0;
                                }
                            }

                            dateCell += 1;
                        }
                    }
                }
            }
            priority -= 1;
        }
    });


    debugger;
    // End Generate resource plan function
}

function getRulesValue(rules) {
    globalRules = rules;
}


function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function removeEmployeeByName(firstName, lastName) {
    var ss = SpreadsheetApp.openByUrl(Sheeturl);
    var ws = ss.getSheetByName("Employee List");
    var data = ws.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) { // Start from 1 to skip header
        if (data[i][0] === firstName && data[i][1] === lastName) { // Assuming first name is in column A and last name in column B
            ws.deleteRow(i + 1); // Adjust for header row (1-based index in sheets)
            return; // Exit after the first match is found and removed
        }
    }
}

function onOpen() {
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var yourNewSheet = activeSpreadsheet.getSheetByName("workLoad Analysis");

    if (yourNewSheet != null) {
        activeSpreadsheet.deleteSheet(yourNewSheet);
    }

    yourNewSheet = activeSpreadsheet.insertSheet();
    yourNewSheet.setName("workLoad Analysis");
}

function createWorkLoadAnalysis() {
    onOpen();
    var ss = SpreadsheetApp.openByUrl(Sheeturl);
    var resourceSheet = ss.getSheetByName('resourcePlan');
    var employeeSheet = ss.getSheetByName('Employee List');
    var createWorkLoadSheet = ss.getSheetByName("workLoad Analysis");
    var emlRow = employeeSheet.getLastRow();
    var relRow = resourceSheet.getLastRow();
    var emName = employeeSheet.getRange(2, 1, emlRow - 1, 1).getValues();
    var resourceSheetFirstRow = resourceSheet.getRange(2, 1, relRow - 1, 1).getValues();
    var resourceSheetThirdRow = resourceSheet.getRange(2, 3, relRow - 1, 1).getValues();
    var emTotalEffortList = [];

    for (i = 0; i < relRow; i++) {
        var effort = resourceSheetThirdRow.flat()[i];
        if (typeof effort == 'number') {
            var totalEffort = resourceSheetThirdRow.flat()[i];
            emTotalEffortList.push(totalEffort);
        }
    }// get total effort list

    var emNamePlacementList = [];
    var relCol = resourceSheet.getLastColumn();

    for (i = 0; i < emName.length; i++) {
        var emNamePlacement = resourceSheetFirstRow.flat().indexOf(emName.flat()[i]);
        emNamePlacementList.push(emNamePlacement);
    }// get name Placement list

    emNamePlacementList.push(relRow - 1);
    var emTotalWorkDay = [];
    var j = 0;
    var i = 0;
    var sum = 0;
    for (i = 0; i < resourceSheetFirstRow.length + 1; i++) {
        if (i == emNamePlacementList[j]) {
            j++;
            emTotalWorkDay.push(sum);
            sum = 0;
        } else {
            var workDay = resourceSheet.getRange(i + 2, 1, 1, relCol).getValues();
            var result = workDay.flat().filter((day) => typeof day === 'number');
            sum = sum + result.length;
        }
    } // get totalWorkDay
    emTotalWorkDay.shift();

    var wlRow = createWorkLoadSheet.getLastRow();
    createWorkLoadSheet.getRange(wlRow + 1, 1, 1, 8).setBackground('#48489c').setFontWeight('bold').setFontColor('white');
    createWorkLoadSheet.getRange(wlRow + 1, 1, 1, 8).setValues([['Employee', 'Expected \nDay \nWorked', 'Days \nWorked', 'Productive \nHrs/Day Goal', 'Productive \nHrs/Day', 'Expected \nTotal \nProductive Hrs', 'Total \nProductive \nHrs', 'User \nCapacity']])

    wlRow = createWorkLoadSheet.getLastRow();
    var expectedDayWork = 2;
    var expectedProductiveHrsPerDay = 2;
    var productiveHrsPerDay = emTotalEffortList.map((emEffort, index) => emEffort / emTotalWorkDay[index]);
    var roundedProductiveHrsPerDay = productiveHrsPerDay.map(hours => Math.round(hours * 10) / 10);
    var expectedTotalProductiveHours = 10;
    var userCapacity = emTotalEffortList.map(emEffort => emEffort / expectedTotalProductiveHours);
    var percentageUserCapacity = userCapacity.map(capacity => capacity * 100);
    var userCapacityWithPercentage = percentageUserCapacity.map(capacity => capacity.toString() + "%");


    for (i = 0; i < emName.length; i++) {
        createWorkLoadSheet.getRange(wlRow + 1 + i, 1, 1, 8).setValues([[emName.flat()[i], expectedDayWork, emTotalWorkDay[i], expectedProductiveHrsPerDay, roundedProductiveHrsPerDay[i], expectedTotalProductiveHours, emTotalEffortList[i], userCapacityWithPercentage[i]]]).setBorder(true, true, true, true, true, false);
    }

    for (i = 0; i < emName.length; i++) {
        if (userCapacity[i] > 1) {
            createWorkLoadSheet.getRange(wlRow + 1 + i, 8).setBackground('red');
        } else if (userCapacity[i] == 1) {
            createWorkLoadSheet.getRange(wlRow + 1 + i, 8).setBackground('green');
        } else {
            createWorkLoadSheet.getRange(wlRow + 1 + i, 8).setBackground('yellow');
        }
    }

}


