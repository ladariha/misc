var currentFile = SpreadsheetApp.getActiveSpreadsheet();
var currentSheet = currentFile.getActiveSheet();

var allStreets = [];
var allDates = [];


String.prototype._glNorm = function () {
    return this.trim().toLowerCase();
};

String.prototype._glIsDate = function () {
    var t = this.match(/^(\d{1,2})\/(\d{4})$/);
    if (t !== null) {
        var month = parseInt(t[1], 10);
        var year = parseInt(t[2], 10);
        if (month > 12 || month < 1) {
            return false;
        }
        if (Math.abs(year - new Date().getUTCFullYear()) > 1) {
            return false;
        }
        return true;
    }
    return false;
};


String.prototype._glIsStreet = function () {
    if (this.toLowerCase() === "dům") {
        return false;
    }
    return this.match(/^\D.*$/) !== null && this.trim().length > 0 ? true : false;
};

function importLeden() {
    new Convertor(0).start();
}
function importUnor() {
    new Convertor(1).start();
}
function importBrezen() {
    new Convertor(2).start();
}
function importDuben() {
    new Convertor(3).start();
}
function importKveten() {
    new Convertor(4).start();
}
function importCerven() {
    new Convertor(5).start();
}
function importCervenec() {
    new Convertor(6).start();
}
function importSrpen() {
    new Convertor(7).start();
}
function importZari() {
    new Convertor(8).start();
}
function importRijen() {
    new Convertor(9).start();
}
function importListopad() {
    new Convertor(10).start();
}
function importProsinec() {
    new Convertor(11).start();
}

function onOpen() {
    var entries = [
        {
            name: "Leden",
            functionName: "importLeden"
        },
        {
            name: "Unor",
            functionName: "importUnor"
        },
        {
            name: "Brezen",
            functionName: "importBrezen"
        },
        {
            name: "Duben",
            functionName: "importDuben"
        },
        {
            name: "Kveten",
            functionName: "importKveten"
        },
        {
            name: "Cerven",
            functionName: "importCerven"
        },
        {
            name: "Cervenec",
            functionName: "importCervenec"
        },
        {
            name: "Srpen",
            functionName: "importSrpen"
        },
        {
            name: "Zari",
            functionName: "importZari"
        },
        {
            name: "Rijen",
            functionName: "importRijen"
        },
        {
            name: "Listopad",
            functionName: "importListopad"
        },
        {
            name: "Prosinec",
            functionName: "importProsinec"
        }
    ];
    SpreadsheetApp.getActiveSpreadsheet().addMenu("GL", entries);
}

function FileData(filename) {
    this.type = null;
    this.data = null;
    this.date = null;
    this.filename = filename;
    this.personName = null;
}

FileData.prototype.getMonth = function () {
    if (this.month) {
        return this.month;
    }

    this.month = parseInt(this.date.split("/")[0], 10);
    return this.month;
};


function FileHandler(file, importedMonth) {

    this.readFile = function () {
        Logger.log("opening file " + file.getName());
        var spreadSheet = SpreadsheetApp.openByUrl(file.getUrl());
        var recapSheet = findSheet(spreadSheet.getSheets());

        if (!recapSheet) {
            SpreadsheetApp.getUi().alert('Nebyl nalezen list se jmenem "rekapitulace" v souboru ' + file.getName());
            return;
        }


        spreadSheet.setActiveSheet(recapSheet);
        var validationReport = validateRecapSheet(spreadSheet, recapSheet);
        Logger.log("Validation for " + file.getName() + " : " + validationReport);
        if (validationReport.length > 0) {
            throw new Error("Problemy se souborem " + file.getName() + ":" + validationReport);
        }
        var d = getData(spreadSheet, recapSheet);
        copySheet(recapSheet, d.personName);
        return d;
    };

    function findSheet(sheets) {
        var pattern = "rekapitulace";
        for (var i = 0, imax = sheets.length; i < imax; i++) {
            if (sheets[i].getName().trim().toLowerCase() === pattern) {
                return sheets[i];
            }
        }
        return null;
    }


    function copySheet(originalSheet, person) {
        var name = (importedMonth + 1) + "/" + person;
        var targetSheet = currentFile.getSheetByName(name);
        if (targetSheet) {
            currentFile.setActiveSheet(targetSheet);
            targetSheet.clear();
        } else {
            currentFile.insertSheet(name);
        }

        var sourceDataRange = originalSheet.getDataRange();
        var sourceSheetValues = sourceDataRange.getValues();

        currentFile.getDataRange().offset(0, 0, sourceDataRange.getNumRows(), sourceDataRange.getNumColumns()).setValues(sourceSheetValues);
        currentFile.setActiveSheet(currentSheet);
    }

    function getDataType(range) {
        return range[3][0].toString()._glNorm();
    }

    function getDataDate(range) {
        var d = range[0][3];
        if (!d || d.toString().length < 6) {
            d = range[0][4].toString();
        }

        return d._glNorm();
    }

    function getDataPersonName(range) {
        return range[2][0].toString();
    }

    function getData(spreadSheet, sheet) {
        var values = sheet.getRange(7, 1, sheet.getLastRow(), 2).getValues();
        var infoRange = sheet.getRange("A1:E7").getValues(); // TODO hardcoded end row

        var data = new FileData(file.getName());
        data.type = getDataType(infoRange);
        data.date = getDataDate(infoRange);
        data.personName = getDataPersonName(infoRange);

        if (allDates.indexOf(data.date) < 0) {
            allDates.push(data.date);
        }

        var stopString = "celkem";
        var stringVal;
        var d = {};
        for (var i = 0, imax = values.length; i < imax; i++) {
            stringVal = values[i][0].toString().trim();

            if (stringVal.toLowerCase() === stopString) {
                break;
            }

            if (stringVal._glIsStreet()) {
                d[stringVal] = parseFloat(values[i][1].toString());
                if (allStreets.indexOf(stringVal) < 0) {
                    allStreets.push(stringVal);
                }
            }
        }

        data.data = d;
        return data;
    }



    function findSheet(sheets) {
        var pattern = "rekapitulace";
        for (var i = 0, imax = sheets.length; i < imax; i++) {
            if (sheets[i].getName().trim().toLowerCase() === pattern) {
                return sheets[i];
            }
        }
        return null;
    }


    function validateRecapSheet(spreadSheet, sheet) {

        var values = sheet.getRange("A1:E7").getValues(); // TODO hardcoded end row
        var validationMsg = "";
        // check table headers
        if (!values[6][0] || values[6][0].toString()._glNorm() !== "dům") {
            validationMsg += "Hlavicka tabulky neodpovida vzoru, bunka A7 obsahuje '" + values[0][0].toString() + "', ale mela by obsahovat 'dům'. ";
        }

        if (!values[6][1] || values[6][1].toString()._glNorm() !== "počet hodin") {
            validationMsg += "Hlavicka tabulky neodpovida vzoru, bunka B7 obsahuje '" + values[0][1].toString() + "', ale mela by obsahovat 'počet hodin'. ";
        }

        // check record type
        if (!values[3][0] || values[3][0].toString()._glNorm().length < 1) {
            validationMsg += "Bunka A4 neobsahuje typ pracovnika.";
        }

        if (!values[3][0] || (values[3][0].toString()._glNorm() !== "technik" && values[3][0].toString()._glNorm() !== "účetní")) {
            validationMsg += "Bunka A4 neobsahuje typ spravny typ pracovnika. Ocekavany typ je 'technik' nebo 'účetní'. ";
        }

        // check name
        if (!values[2][0] || values[2][0].toString()._glNorm().length < 1) {
            validationMsg += "Bunka A3 neobsahuje zadny text, ale mela by obsahovat jmeno.";
        }

        // check date
        var hasDate = true;
        if ((!values[0][3] || !values[0][3].toString()._glIsDate()) && (!values[0][4] || !values[0][4].toString()._glIsDate())) {
            validationMsg += "Bunka D1 ani bunka E1 neobsahuje datum ve formatu MM/YYYY, napr. '10/2014'.";
            hasDate = false;
        }

        var m = null;
        if (hasDate) {

            if (values[0][3]) {
                m = values[0][3].toString()._glNorm().split("/")[0];
            } else if (values[0][4]) {
                m = values[0][4].toString()._glNorm().split("/")[0];
            }

            if (parseInt(m, 10) !== (importedMonth + 1)) {
                validationMsg += "Soubor obsahuje datum pro mesic " + m + ", ale pozadovany mesic pro import dat je " + (importedMonth + 1) + ". ";
            }

        }

        return validationMsg;
    }
}


function Convertor(importedMonth) {

    var self = this;
    var sourceFolder = null;

    this.start = function () {

        sourceFolder = currentFile.getSheetByName("nastaveni").getRange("B3").getValue();
        if (!sourceFolder) {
            throw new Error('Prosim vloz do bunky B3 v listu "nastaveni" nazev hlavni slozky');
        }


        try {

            var files = findFiles(sourceFolder);

            if (files.length === 0) {
                throw new Error('Nebyly nalezeny zadne odpovidajici soubory');
            }

            var allData = [];
            for (var i = 0, imax = files.length; i < imax; i++) {
                allData.push(new FileHandler(files[i], importedMonth).readFile());
            }

            new Printer(allData, importedMonth).print();
        } catch (err) {
            Logger.log(err.stack);
            Logger.log(err);
            SpreadsheetApp.getUi().alert(err);
        }
    };

    function findFiles(folder) {
        Logger.log("finding files in " + folder + "/" + importedMonth);
        var folders = DriveApp.getFoldersByName(folder);
        var count = 0;
        var targetFolder = null;
        var _f;
        while (folders.hasNext()) {
            count++;
            _f = folders.next();
            targetFolder = targetFolder ? targetFolder : _f;
        }

        if (count !== 1) {
            throw new Error('Nalezeno ' + count + ' slozek stejneho jmena, prosim pouzij unikatni jmeno existujici slozky');
        }


        var monthFolders = targetFolder.getFoldersByName((importedMonth + 1));
        var targetMonth = null;
        count = 0;
        while (monthFolders.hasNext()) {
            count++;
            _f = monthFolders.next();
            targetMonth = targetMonth ? targetMonth : _f;
        }

        if (count !== 1) {
            throw new Error('Nalezeno ' + count + ' slozek stejneho jmena "' + (importedMonth + 1) + '" pro vybrany mesic');
        }



        var files = targetMonth.getFiles();
        var _files = [];
        while (files.hasNext()) {
            var f = files.next();
            if (f.getMimeType() === "application/vnd.google-apps.spreadsheet") {
                _files.push(f);
                Logger.log("Adding " + f.getName() + " to the list");
            }
        }

        return _files;
    }

}

function Printer(data, importedMonth) {

    var price = parseInt(currentFile.getSheetByName("nastaveni").getRange("B2").getValue(), 10);
    var sheet = null;
    var existingStreets = {};
    var monthMap = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"];


    this.print = function () {
        sheet = currentFile.getSheetByName("Výtěžnost");
        var s = sheet;
        if (sheet) {
            currentFile.setActiveSheet(sheet);
        } else {
            sheet = currentFile.insertSheet("Výtěžnost");
        }
        getExistingStreets();
        var offset = 0;
        for (var i = 0, imax = allStreets.length; i < imax; i++) {
            if (existingStreets.hasOwnProperty(allStreets[i])) {
                // find and update data :( 
                updateStreet(allStreets[i]);
            } else {
                // find suitable row where to start with inserting
                offset = sheet.getLastRow() + 2;
                printStreet(allStreets[i], offset);
            }
        }
    };


    function updateStreet(street) {
        var startRow = existingStreets[street]; // 0 is A1
        var monthColumn = sheet.getRange(monthMap[importedMonth] + (6 + startRow) + ":" + monthMap[importedMonth] + (9 + startRow));


        Logger.log("updating street " + street);
        var relevantRecords = [];

        for (var i = 0, imax = data.length; i < imax; i++) { // TODO add condition for month as well!!!!
            if (data[i].data.hasOwnProperty(street)) { // process just fitting records
                relevantRecords.push(data[i]);
            }
        }

        if (relevantRecords.length < 1) {
            throw new Error("Nebyly nalezeny vykazy pro ulici " + street);
        }
        var r = collectDataByDate(relevantRecords, street);
        var values = [
            [r.sums["technik"][importedMonth]],
            [r.sums["účetní"][importedMonth]],
            [r.sums["technik"][importedMonth] + r.sums["účetní"][importedMonth]],
            [price * (r.sums["technik"][importedMonth] + r.sums["účetní"][importedMonth])]
        ];
        monthColumn.setValues(values);
    }

    function getExistingStreets() {
        var lastRow = sheet.getLastRow();
        if (lastRow === 0) {
            return;
        }
        var range = sheet.getRange("A1:A" + lastRow);
        var values = range.getValues();
        var _s;
        for (var i = 0, imax = values.length; i < imax; i++) {
            _s = values[i][0].toString();
            if (_s.length > 0) {
                existingStreets[_s] = i;
            }
        }
    }

    function collectDataByDate(records, street) {

        var collection = {
            technician: "",
            accountant: "",
            sums: {
                "technik": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                "účetní": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
            }
        };

        var _month;
        var alreadyProcessed = {
            "technik": [],
            "účetní": []
        };
        for (var i = 0, imax = records.length; i < imax; i++) {
            _month = records[i].getMonth() - 1;
            if (records[i].type === "technik") {
                collection.technician = records[i].personName;
            } else {
                collection.accountant = records[i].personName;
            }
            if (alreadyProcessed[records[i].type].indexOf(_month) > -1) {
                throw new Error("Byly nalezeny vice nez 2 vykazy pro ulici " + street + " v mesici " + (_month + 1) + " s pracovnikem typu " + records[i].type);
            } else {
                alreadyProcessed[records[i].type].push(_month);
            }
            collection.sums[records[i].type][_month] = records[i].data[street];
        }

        return collection;
    }


    function printStreet(street, offset) {
        Logger.log("processing street " + street);
        var relevantRecords = [];

        for (var i = 0, imax = data.length; i < imax; i++) {
            if (data[i].data.hasOwnProperty(street)) { // process just fitting records
                relevantRecords.push(data[i]);
            }
        }

        if (relevantRecords.length < 1) {
            throw new Error("Nebyly nalezeny vykazy pro ulici " + street);
        }

        createStreetTable(street, offset, collectDataByDate(relevantRecords, street));
    }

    function createStreetTable(street, offset, config) {

        var range = sheet.getRange("A" + (2 + offset) + ":O" + (11 + offset));
        var t = config.sums["technik"];
        var u = config.sums["účetní"];

        for (var i = 0, imax = 12; i < imax; i++) {
            t[12] += t[i];
            u[12] += u[i];
        }

        var values = [
            [street, "správa", "", "technik", config.technician, "", "", "KČ/hod", "", "", "", "", "", "", ""],
            ["", "úklid", "", "účetní", config.accountant, "", "", price, "", "", "", "", "", "", ""],
            ["", "údržba", "", "", "", "", "", "", "", "", "", "", "", "", ""],
            ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
            ["", "", "Leden", "Únor", "Březen", "Duben", "Květen", "Červen", "Červenec", "Srpen", "Září", "Říjen", "Listopad", "Prosinec", "celkem"],
            ["", "technik", t[0], t[1], t[2], t[3], t[4], t[5], t[6], t[7], t[8], t[9], t[10], t[11], t[12]],
            ["", "účetní", u[0], u[1], u[2], u[3], u[4], u[5], u[6], u[7], u[8], u[9], u[10], u[11], u[12]],
            ["", "celkem", t[0] + u[0], t[1] + u[1], t[2] + u[2], t[3] + u[3], t[4] + u[4], t[5] + u[5], t[6] + u[6], t[7] + u[7], t[8] + u[8], t[9] + u[9], t[10] + u[10], t[11] + u[11], t[12] + u[12]],
            ["", "náklad", price * (t[0] + u[0]), price * (t[1] + u[1]), price * (t[2] + u[2]), price * (t[3] + u[3]), price * (t[4] + u[4]), price * (t[5] + u[5]), price * (t[6] + u[6]), price * (t[7] + u[7]), price * (t[8] + u[8]), price * (t[9] + u[9]), price * (t[10] + u[10]), price * (t[11] + u[11]), price * (t[12] + u[12])],
            ["", "rozdíl", "", "", "", "", "", "", "", "", "", "", "", "", ""]
        ];
        range.setValues(values).setBorder(true, true, true, true, false, false);

        sheet.getRange("B" + (6 + offset) + ":O" + (6 + offset)).setBorder(false, false, true, true, false, false);
        sheet.getRange("B" + (9 + offset) + ":O" + (9 + offset)).setBorder(false, false, true, true, false, false);
        addFormulas(sheet, range, offset);
    }

    function addFormulas(sheet, range, offset) {

        var hoursTotalFormulas = [];
        var monthExpensesFormulas = [];
        var diffsFormulas = [];
        for (var i = 0, imax = monthMap.length; i < imax; i++) {
            // month total
            hoursTotalFormulas.push(formulas.getHoursTotal(offset, monthMap[i]));
            monthExpensesFormulas.push(formulas.getMonthExpense(offset, monthMap[i]));
            diffsFormulas.push(formulas.getMonthDiff(offset, monthMap[i]));
        }
        hoursTotalFormulas.push(formulas.getHoursTotalYear(offset));
        monthExpensesFormulas.push(formulas.getTotalExpense(offset));
        diffsFormulas.push(formulas.getTotalDiff(offset));

        sheet.getRange("C" + (9 + offset) + ":O" + (9 + offset)).setFormulas([hoursTotalFormulas]);
        sheet.getRange("C" + (10 + offset) + ":O" + (10 + offset)).setFormulas([monthExpensesFormulas]);
        sheet.getRange("C" + (11 + offset) + ":O" + (11 + offset)).setFormulas([diffsFormulas]);
    }
}

var formulas = {
    getHoursTotal: function (offset, monthColumn) {
        return "=SUM(" + monthColumn + (7 + offset) + ":" + monthColumn + (8 + offset) + ")";
    },
    getHoursTotalYear: function (offset) {
        return "=SUM(C" + (9 + offset) + ":N" + (9 + offset) + ")";
    },
    getMonthExpense: function (offset, monthColumn) {
        return "=" + monthColumn + (9 + offset) + "*H" + (offset + 3);
    },
    getTotalExpense: function (offset) {
        return "=SUM(C" + (10 + offset) + ":N" + (10 + offset) + ")";
    },
    getMonthDiff: function (offset, monthColumn) {
        return "(C" + (offset + 2) + "+C" + (offset + 4) + ")-" + monthColumn + (offset + 10);
    },
    getTotalDiff: function (offset) {
        return "=SUM(C" + (11 + offset) + ":N" + (11 + offset) + ")";
    }
};
