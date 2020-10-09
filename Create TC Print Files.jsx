/*

Note: This is specifically a script written for Adobe Indesign CS6 in JavaScript. 
It was used to automate a monthly periodical imposition named "TC". It relinked the final InDesign file to templates and exported the production files.  
It also parsed a labels csv, and created the addressed envelopes & addressed body files. 
(Note: InDesign uses its own flavor of javascript that can read / write files) 

*/

var months = ['01-January', '02-February', '03-March', '04-April', '05-May', '06-June', '07-July', '08-August', '09-September', '10-October', '11-November', '12-December']; // list of months
var myPDFExportPreset = app.pdfExportPresets.item("High Quality Print(1200)"); // Change myPDFpreset to the name of the preset you want to use
var nbs = '\xa0\xa0\xa0\xa0\xa0\xa0\xa0\xa0\xa0\xa0\xa0\xa0\xa0\xa0\xa0\xa0\xa0\xa0\xa0\xa0\xa0\xa0\xa0\xa0\xa0\xa0' // non breaking space, used in the GUI window

// template file locations
var continentalTCBody = "X:\\Production\\Periodicals\\TC\\Imposition\\Continental TC Body.indt";
var foreignTCBody = "X:\\Production\\Periodicals\\TC\\Imposition\\Foreign TC Booklet.indt";
var bothTCCovers = "X:\\Production\\Periodicals\\TC\\Imposition\\Both TC Cover 2up.indt";
var tcEnvelopeImposition = "X:\\Production\\Periodicals\\TC\\Imposition\\TC Envelopes.indt";
var tcEnvelopeSrcCsv = "\\\\CCC-DB\\DB\\CCC\\DF\\TC-FOREIGN-SINGLE.CSV"
var tcAddressedSrcCsv = "\\\\CCC-DB\\DB\\CCC\\DF\\TC-US-SINGLE.CSV"
var productionFolder = "X:\\Production\\Periodicals\\TC";

var cmsFolder = 'X:\\CMS\\25'; // project folder path
var januaryCMSFolder = 25686; //the january of the current printing year
var folderSeparator = "\\" // escaped backslash string

var curYear = null;
var curMo = null;
var tcFile = null;
var tcDoc = null;
var pBar = new ProgressBar("Script Title", 2); //, 18);
var addressedPageRange = null;


var okToProceed = false; // if user cancels out of the window, main function is canceled

//Suppress all dialogs
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.neverInteract;

//getCurrentTC();

createWindow();

if (okToProceed) { //if the create window is successful in picking a file.
    try {

        // updates any links that are out of date.
        tcDoc = app.open(tcFile, false);
        updateLinks(tcDoc, true);
        tcDoc.close();

        //alerts if labels were modified more than a day ago.
        var tempFile = new File(tcEnvelopeSrcCsv); //temp file is just a way to see the last modified date
        var result = (Date.now() - tempFile.modified) / 3600000; //gives the hours difference
        if (result > 24 || result < -24) {
            alert('The labels CSV was modified ' + (result / 24).toFixed(0) + ' day(s) ago. \n\nPlease make sure that the periodical labels have been created before clicking OK.');
        }
        var startTime = Date.now();

        //not important, but opened in specific order so that they show on screen when being worked on. ie. active document
        var continentalTCBodyDoc = app.open(continentalTCBody, false);
        var bothTCCoversDoc = app.open(bothTCCovers, false);
        var foreignTCBodyDoc = app.open(foreignTCBody, false);

        //configure steps, loads progress bar

        var addressesCt = (loadCSV(tcEnvelopeSrcCsv).length + loadCSV(tcAddressedSrcCsv).length) - 2;
        var steps = (continentalTCBodyDoc.links.count() + foreignTCBodyDoc.links.count() + bothTCCoversDoc.links.count() + addressesCt);
        steps += (4 - 2) //for exportation of pdf files, - 2 for off by 1 errors
        steps += 2 // for tc envelope creation steps
        pBar = new ProgressBar('Script Title', steps)
        pBar.show();

        //TC Envelope imposition
        pBar.updateStatus("Creating TC Envelopes");
        imposeTcEnvelope();

        //foreign TC Body Relink & Export

        relinkFiles(tcFile, foreignTCBodyDoc, pBar);
        exportPDF(foreignTCBodyDoc, " TC Overseas Body.pdf", myPDFExportPreset, null);
        foreignTCBodyDoc.close(SaveOptions.no);
        pBar.performStep();

        //TC Covers Relink & Export

        relinkFiles(tcFile, bothTCCoversDoc, pBar);
        exportPDF(bothTCCoversDoc, " TC Continental Cover 2up.pdf", myPDFExportPreset, "1");
        pBar.performStep();
        exportPDF(bothTCCoversDoc, " TC Overseas Cover 2up.pdf", myPDFExportPreset, "2")
        pBar.performStep();
        bothTCCoversDoc.close(SaveOptions.no);

        //Continental TC Body Relink & Export

        relinkFiles(tcFile, continentalTCBodyDoc, pBar);
        exportPDF(continentalTCBodyDoc, " TC Continental Body.pdf", myPDFExportPreset, null);
        //continentalTCBodyDoc.print(false, 'TC-Print');
        //continentalTCBodyDoc.save();


        //Continental TC Body - Addressed versions export
        continentalTCBodyDoc.windows.add();
        imposeTcAddressed(continentalTCBodyDoc);
        continentalTCBodyDoc.close(SaveOptions.no);
        alert("Finished!\nIt took " + ((Date.now() - startTime) / 60000).toFixed(2) + " minutes to process.")

        pBar.performStep();
        pBar.updateStatus("Done!")
        $.sleep(1000); // shows status for a second.
        //copyToClipboard(getTCProductionFolder() + folderSeparator + curMo + " TC Continental Body-Printed.pdf");
        //alert('Please print pages 1-19 of the TC document to the production folder. \n(Paste, Copy before printing, paste in copied file name when printing...)');

        var myFold = new Folder(getTCProductionFolder());
        myFold.execute(); //opens tc production folder
    } catch (err) {
        alert(err)
    }

}


//close pBar
pBar.hide();

//Restore dialogs
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;

//Restore page range
app.pdfExportPreferences.pageRange = PageRange.ALL_PAGES;



/*
 * Creates the Addressed TC files. 
 */
function imposeTcAddressed(tcDoc) {
    app.activeWindow.transformReferencePoint = AnchorPoint.topLeftAnchor;
    var pStyle = null;
    var numStyle = null;

    try {
        pStyle = tcDoc.paragraphStyles.add('tcAddresses');
        with(pStyle) {
            appliedFont = "Arial";
            pointSize = 12;
        }
    } catch (e) {
        pStyle = tcDoc.paragraphStyles.item("tcAddresses");
    }

    try {
        numStyle = tcDoc.paragraphStyles.add('numStyle');
        with(numStyle) {
            appliedFont = "Arial";
            pointSize = 7;
        }
    } catch (e) {
        numStyle = tcDoc.paragraphStyles.item("numStyle");
    }

    var myPg = tcDoc.pages[19];
    var numTxt = myPg.textFrames.add();
    numTxt.absoluteRotationAngle = 90;
    numTxt.geometricBounds = [3.5, 2.5, 4, 2.75]; // [ y1, x1, y2, x2] y1 & x1 = top left, y2 & x2 = bottom right
    var myTxt = myPg.textFrames.add();
    myTxt.absoluteRotationAngle = 90;
    myTxt.geometricBounds = [4.25, 1.25, 8.125, 2.75]; // [ y1, x1, y2, x2]
    var addresses = loadCSV(tcAddressedSrcCsv);
    for (var i = 0; i < addresses.length; i++) { //goes through each address & exports a pdf of each
        var curPadTC = ("0000" + (i + 1)).slice(-4) // gets the padded current number. ie. if i == 0, curPadTC == '0001'
        pBar.updateStatus("Imposing Addressed TC " + i + " of " + addresses.length);
        numTxt.contents = curPadTC;
        numTxt.parentStory.texts.item(0).applyParagraphStyle(numStyle, true);
        myTxt.contents = addresses[i];
        myTxt.parentStory.texts.item(0).applyParagraphStyle(pStyle, true);
        //function exportPDF(impositionTemplateFile, fileName, exportPreferences, pageRange){
        var custFp = getTCProductionFolder() + folderSeparator + "Bodies" + folderSeparator + curPadTC + ".pdf";
        exportPDF(tcDoc, " Body - " + curPadTC + ".pdf", myPDFExportPreset, addressedPageRange, custFp);
        pBar.performStep();
    }

}



function imposeTcEnvelope() {
    var doc = app.open(File(tcEnvelopeImposition));
    app.activeWindow.transformReferencePoint = AnchorPoint.topLeftAnchor;
    var pStyle = null
    try {
        pStyle = doc.paragraphStyles.add('tcAddresses');
        with(pStyle) {
            appliedFont = "Arial";
            pointSize = 12;
        }
    } catch (e) {
        pStyle = doc.paragraphStyles.item("tcAddresses");
        //alert("tcAddresses paragraph style already exists?");
    }

    var addresses = loadCSV(tcEnvelopeSrcCsv);
    for (var i = 0; i < addresses.length; i++) { //goes through each address
        pBar.updateStatus("Imposing Envelope " + i + " of " + addresses.length);
        if (addresses[i] != "") { //not a blank address
            var myPg = null;
            if (i == 0) { //sets correct page to be current page.
                myPg = doc.pages[0];
            } else {
                myPg = doc.pages.add();
            }

            var myTxt = myPg.textFrames.add(); //adds the address text frame & sets correct sizing
            myTxt.geometricBounds = [3, 4.0, 5.5, 8.5]; // [ y1, x1, y2, x2]
            myTxt.parentStory.texts.item(0).applyParagraphStyle(pStyle, true);
            //alert(typeof addresses[i]);
            myTxt.contents = addresses[i]; //TextFrameContents.placeholderText;	
            if (addresses[i].toUpperCase().indexOf("EGYPT") >= 0) {
                //apply master b page
                myPg.appliedMaster = doc.masterSpreads.item('B-Master');
            } else {
                //apply master a page
                myPg.appliedMaster = doc.masterSpreads.item('A-Master');
            }
        }

        pBar.performStep();
    }
    exportPDF(doc, " TC Envelopes.pdf", myPDFExportPreset, null);
    var FN = getTCProductionFolder() + folderSeparator + curMo + " TC Envelopes.indd";
    doc.save(File(FN), true);
    doc.close(SaveOptions.no)


}


/**
 * gets the addresses from the specified file. Returns array, or null if unable to load
 * @param {fp} the string file path to load
 */
function loadCSV(fp) {
    // loads the file
    var results = [];
    fp = new File(fp);
    //results.push(fp.modified);
    fp.open("r");
    var strLine = fp.readln();
    var i = 0;
    while (strLine != "" && i < 10000) {
        if (i == 0) strLine = fp.readln(); //skips first line

        var myArray = splitCSV(strLine);

        if (myArray != null) {
            //alert(myArray.join('\n'));
            results.push(myArray.join('\n') + '\n');
        }

        strLine = fp.readln();
        i++;
    }


    return results;

}

/**
 * parses 1 line from csv, removes quote marks and cleans it up. Returns array of string
 */
function splitCSV(csvStr) {
    var matches = csvStr.match(/(\s*"[^"]+"\s*|\s*[^,]+|,)(?=,|$)/g);
    var textMatches = [];
    for (var n = 0; n < matches.length; ++n) {
        var match = matches[n];
        //matches[n] = matches[n].trim();

        if (match == ',' || match == ' ') match = '';
        match = match.replace('"', ""); //removes quote mark
        match = match.replace('"', ""); //removes quote mark
        match = match.replace(/^\s+/, '').replace(/\s+$/, ''); //this is a trim function
        if (n == 8 || n == 9 || n == 10) {
            match = ""; //removes 'new container' & misc marks
        }
        matches[n] = match;
        if (match != '') textMatches.push(match);

    }

    if (this[0] == ',') matches.unshift("");
    //return matches;
    return textMatches;
}

/**
 * copies text to the indesign clipboard NOT the system clipboard. (trying to paste while indesign is closed / busy is problematic)
 * @param {text} the text to copy to the clipboard
 */
function copyToClipboard(text) {
    var myDoc = app.documents.add(false);
    var myWindow = myDoc.windows.add();
    with(myDoc.documentPreferences) {
        pageHeight = '11in';
        pageWidth = '8.5in'
        pageOrientation = PageOrientation.portrait;
        pagesPerDocument = 1;
    }
    var myTextFrame = myDoc.pages[0].textFrames.add(); //app.activeWindow.activePage.textFrames.add();
    myTextFrame.geometricBounds = [.5, .5, 10.5, 8]
    myTextFrame.contents = text; //TextFrameContents.placeholderText;

    app.selection = myTextFrame.texts;
    app.copy();
    app.paste();
    myDoc.close(SaveOptions.no);
    $.sleep(200);
}

/**
 * returns the selected TC file
 * @param {cmsNumber} the cms folder number that contains the TC
 */
function getCurrentTC(cmsNumber) {
    //This section shows the cms folder & returns the final TC .indd
    //var result = prompt("What is the current CMS folder #?", "24842")

    var cmsTCFolder = new Folder(cmsFolder + folderSeparator + cmsNumber);
    tcFile = cmsTCFolder.openDlg("Select Files:", "*final.indd,*.indd", false);
    return tcFile;
    // tcDoc = app.open(tcFile);

}



/**
 * updates *all* links in document
 * @param {document} doc The document to update links on.
 * @param {boolean} saveDoc Set to true if wanting to save after updating links 
 */
function updateLinks(doc, saveDoc) {
    if (doc != null) {
        for (i = 0; i < doc.links.count(); i++) {
            try {
                doc.links[i].update();
            } catch (err) {
                //alert(i + "\n" + err);
                // usually happens when link can't be found.
            }
        }
        if (saveDoc) {
            if (doc.modified == true) {
                var saved = false;
                try {
                    doc.save();
                    saved = true;
                } catch (err) {}
                if (!saved) {
                    doc.close(SaveOptions.no)
                    throw "Not able to save file! \n(Probably read only)"

                }

            }

        }

    }
}

/**
 * Relinks the imposition template file pages to the current tc file pages 
 * @param {appDocument} curTcFile The master tc file
 * @param {appDocumentTemplate} impositionTemplateFile The imposition file 
 * @param {ProgressBar} pBar Progress bar
 */
function relinkFiles(curTcFile, impositionTemplateFile, pBar) {
    //alert("here. i.. am...")
    //pBar = new ProgressBar("Title", 5, )
    var linkCt = impositionTemplateFile.allGraphics.length;
    //    alert(linkCt);
    for (i = 0; i < linkCt; i++) {
        pBar.updateStatus("Relinking page " + (i + 1) + " of " + linkCt);
        impositionTemplateFile.allGraphics[i].itemLink.relink(curTcFile);
        pBar.performStep();
    }

}


/**
 * Exports the impositionTemplateFile to the production folder with the specified filename, exportPreferences & pageRange. If custFullFilePath is specified, overrides fileName.
 */
function exportPDF(impositionTemplateFile, fileName, exportPreferences, pageRange, custFullFilePath) {
    var fn = curMo + fileName;
    pBar.updateStatus('Exporting ' + fn);
    switch (pageRange) {
        case null:
            //fall through to next
        case 'all':
            app.pdfExportPreferences.pageRange = PageRange.ALL_PAGES;
            break;
        default:
            app.pdfExportPreferences.pageRange = pageRange;
    }

    var fullFilePath = getTCProductionFolder() + folderSeparator + curMo + fileName;
    if (custFullFilePath != null) {
        fullFilePath = custFullFilePath;
        //alert(fullFilePath);
    }

    impositionTemplateFile.exportFile(
        ExportFormat.pdfType,
        File(fullFilePath),
        false,
        exportPreferences);
}

/**
 * Returns the path to the correct month's production folder. It creates the folder if it doesn't exist.
 */
function getTCProductionFolder() {

    var newProdFolder = productionFolder + folderSeparator + curYear + folderSeparator + curMo
    var f = new Folder(newProdFolder);
    if (!f.exists) { //creates production folder if it doesn't exist
        f.create();
    }
    var bodiesFolder = newProdFolder + folderSeparator + 'Bodies';
    var b = new Folder(bodiesFolder);
    if (!b.exists) { //creates bodies folder if it doesn't exist
        b.create();
    }
    return newProdFolder;

}


/**
 * returns the 15th of the current month + monthsToAdd
 * @param {number} monthsToAdd number of months to add
 * @return {date}
 */
function addMonths(monthsToAdd) {
    var t = new Date();
    var d = new Date(t.getFullYear(), t.getMonth(), 15);
    d.setMonth(d.getMonth() + monthsToAdd);
    return d;
}

/**
 * Returns the corresponding cms number
 */
function getCMSFolderNumber(printMonth) {
    return januaryCMSFolder + printMonth
}


/**
 * Creates a GUI window that allows the user to specify TC file & TC year
 */
function createWindow() {
    //creates the date for '2' months into the future
    var d = addMonths(2)
    //creates a new window
    var box = new Window('dialog', "Export: The Christian");

    box.orientation = 'column';
    box.alignChildren = 'left';

    // text for instructions
    box.add('statictext', undefined, 'Make sure year and month are correct! ')
    box.add('statictext', undefined, 'The selected TC file needs to NOT be read only.')
    box.add('statictext', undefined, '')


    //panel 1 - box layout
    var panel1 = box.add('panel', undefined, "The Christian - Specs");
    panel1.alignChildren = 'left';

    var group1 = panel1.add('group', undefined);
    group1.orientation = 'row';

    group1.add('statictext', undefined, 'TC Year:');
    var tcYearTxt = group1.add('edittext', undefined, d.getFullYear());

    var group2 = panel1.add('group', undefined);
    group2.orientation = 'row';
    group2.add('statictext', undefined, 'TC Month:');
    var tcMonthlist = group2.add('dropdownlist', undefined, undefined, {
        items: months
    });

    //selects 'current' month
    tcMonthlist.selection = d.getMonth();

    //third row, first panel
    var group3 = panel1.add('group', undefined);
    group3.orientation = 'row';
    group3.add('statictext', undefined, 'CMS Folder:');
    var cmsNumberTxt = group3.add('edittext', undefined, getCMSFolderNumber(d.getMonth())); //box.panel1.add

    //select the file to process
    var group4 = panel1.add('group', undefined)
    group4.orientation = 'row'
    group4.add('statictext', undefined, 'TC File:');
    var tcFileTxt = group4.add('statictext', undefined, 'None Selected' + nbs);
    var fileBtn = panel1.add('button', undefined, 'Choose File...', {
        name: 'PickFile'
    });

    fileBtn.onClick = function () {
        //opens the open file dialog. 
        try {
            tcFile = getCurrentTC(cmsNumberTxt.text)
            if (tcFile == null) {
                tcFileTxt.text = 'None Selected' + nbs; //sets the text to show 'none selected'.
            } else {
                tcFileTxt.text = tcFile.displayName; //sets the text to show the file name.
            }
        } catch (err) {
            alert(err);
        }

    }

    // next group
    var group5 = panel1.add('group', undefined);
    group5.orientation = 'row';
    var rdbSinglePg = group5.add('radiobutton');
    rdbSinglePg.text = 'Address Only (5-6 minutes)';
    var rdbAllPg = group5.add('radiobutton');
    rdbAllPg.text = 'All pages (25-30 minutes)';
    rdbAllPg.value = true;

    panel1.add('statictext', undefined, '') //spacing after settings

    var group6 = panel1.add('group', undefined);
    group6.orientation = 'row';
    group6.alignment = 'right'

    var btnCreate = group6.add('button', undefined, 'Create Print Files', {
        name: 'btnCreate'
    });
    var btnCancel = group6.add('button', undefined, 'Cancel', {
        name: 'btnCancel'
    });

    btnCreate.onClick = function () {
        //processes options first, then if ok, closes it.
        //final ok button
        if (tcFile != null) {
            if (rdbAllPg.value == true) {
                // all pages selected to export
                addressedPageRange = 'all'
            } else {
                addressedPageRange = '20'
            }
            curYear = tcYearTxt.text;
            curMo = tcMonthlist.selection.text;
            okToProceed = true;
            box.hide();
        } else {
            alert("Please choose a file first!")
        }
    }
    btnCancel.onClick = function () {
        //cancels the operation
        box.close();

    }
    box.show()
}


/**
 * Creates a progress window with 2 progress bars
 * @param {string} title The title of the progressbar.
 * @param {number} maxValue The number of steps (25?).
 */
function ProgressBar(title, maxValue) { //, myLinksValue) {
    var title = 'Relinking TC Progress';
    var w = new Window('palette', ' ' + title, {
            x: 0,
            y: 0,
            width: 600,
            height: 84
        }),
        pBar = w.add('progressbar', {
            x: 20,
            y: 12,
            width: 550,
            height: 12
        }, 1, maxValue),
        status = w.add('statictext', {
            x: 20,
            y: 60,
            width: 550,
            height: 20
        }, '');
    //status.justify = 'center';
    w.center();
    this.performStep = function () {
        pBar.value++;
    };
    this.updateStatus = function (msg) {
        status.text = msg;
    };
    this.hide = function () {
        w.hide();
    };
    this.show = function () {
        w.show();
    };
};
