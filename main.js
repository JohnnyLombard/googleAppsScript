var EVENT_IMPORTED = "AJOUTE"; // Ajoutera le texte "AJOUTE" dans la colonne J
var ss = SpreadsheetApp.getActiveSpreadsheet();
var ActiveSheet = SpreadsheetApp.openById("1LJp4Ecau5UCK_aCA9vZiofR1__dLLl66P-6OjnPjIYM").getSheetByName("Réponses au formulaire 1");

function onOpen() {
    var menuEntries = [{ name: "Ajouter les événements à l'agenda", functionName: "importCalendar" }];
    ss.addMenu("Agenda", menuEntries); // Pour ajouter une menu Agenda et un sous-menu "ajouter les événements" afin de tester manuellement le programme
}

//--------------------------------------------------------------------------------------------------
//----------------------------------------ARCHIVAGE-------------------------------------------------OK
//--------------------------------------------------------------------------------------------------

function onSelectionChange(e) {

    var range = e.range;
    var spreadSheet = e.source;
    var sheetName = spreadSheet.getActiveSheet().getName();
    var column = range.getColumn();
    var row = range.getRow();
    var dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

    if (row > 1 && column === 12 && sheetName === "Réponses au formulaire 1") {
        if (dataSheet.getRange(row, 13).getValue() === "x") {
            dataSheet.getRange(row, 13).setValue('');
        } else {
            dataSheet.getRange(row, 13).setValue('x');
        }
    }

}

function copyRows() {

    var copySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Archive");
    var dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Réponses au formulaire 1");
    var dataRange = dataSheet.getDataRange();
    var dataValues = dataRange.getValues();

    for (var i = 1; i < dataValues.length; i++) {
        if (dataValues[i][12] === "x") {
            copySheet.appendRow([dataValues[i][0],
            dataValues[i][1],
            dataValues[i][2],
            dataValues[i][3],
            dataValues[i][4],
            dataValues[i][5],
            dataValues[i][6],
            dataValues[i][7],
            dataValues[i][8],
            dataValues[i][9],
            dataValues[i][10],
            dataValues[i][11]]);
        }
    }

    for (var i = 1; i < dataValues.length; i++) {
        if (dataValues[i][12] === "x") {
            var clearRow = i + 1;
            dataSheet.getRange('A' + clearRow + ':M' + clearRow).clear();
        }
    }
}

function addMenu() {
    var menu = SpreadsheetApp.getUi().createMenu('Custom');
    menu.addItem('Copy Rows', 'copyRows');
    menu.addToUi();
}

// function onOpen(e)
// {
//   addMenu(); 
// }

function compare(x, y) {        //--Fonction de comparaison si pas de disponibilité
    return x - y;
}

//----------------------------------------------------------------------------------------------------------
//--------------------------------RESERVATION---------------------------------------------------------------
//----------------------------------------------------------------------------------------------------------

function importCalendar() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var startcolumn = 2;  // Première colonne de prise en compte des données, soit la colonne B 
    var numcolumns = 30;  // Nombre de colonne
    var dataRange = sheet.getRange(startcolumn, 1, numcolumns, 13)   // Nombre de colonne contenant des données
    var data = dataRange.getValues();

    for (var i = 0; i < data.length; i++) {
        //var currentRow = data[i];
        var date01 = new Date();
        var column = data[i];
        var DateDebut = column[1];        // Colonne B - date de début réservation
        var DateFin = column[2];          // Colonne C - date de fin réservation
        var Nom = column[3];              // Colonne D - Nom
        var Prenom = column[4];           // Colonne E - prénom
        var Mail = column[5];             // Colonne F - mail du commercial
        var Distance = column[6];         // Colonne G - Km à éffectuer
        var ServiceInfo = column[7];      // Colonne H - Service Info
        var Vehicule = column[8];         // Colonne I - Véhicule

        //------------------------------------------------------------------------------------------------
        //--------------------------VERIFIER LES RESERVATION TERMINE--------------------------------------OK
        //------------------------------------------------------------------------------------------------


        if (date01.valueOf() > DateFin && DateDebut != "") {
            sheet.getRange(startcolumn + i, 11).setValue("TERMINE");
            sheet.getRange(startcolumn + i, 13).setValue("x")
        }


        //------------------------------------------------------------------------------------------------
        //---------------------------réservation voiture Service INFO-------------------------------------
        //------------------------------------------------------------------------------------------------

        if (ServiceInfo == "Oui" && Vehicule == "") {
            sheet.getRange(startcolumn + i, 9).setValue("Kangoo");

            for (var b = 0; b < data.length; b++) {
              currentRow = data[b];
              var sdate = currentRow[1];
              var edate = currentRow[2];
              if((DateDebut >= sdate && DateDebut<= edate && currentRow[8] == "Kangoo" && currentRow[7] == "Non" ) || (DateFin >= sdate && DateFin <= edate && currentRow[8] == "Kangoo" && currentRow[7] == "Non" )) {
                sheet.getRange(startcolumn + b, 13).setValue("x");
                sheet.getRange(startcolumn + b, 12).setValue("Oui");
                sheet.getRange(startcolumn + b, 11).setValue("TERMINE");

                var sendTo = data[b][5];
       	        var messageConfirme = "";
                                        messageConfirme += "<p>Bonjour, " +data[b][3] + " </p>" +
                                        "<p>Nous vous informons de l'annulation du Kangoo, car le service Informatique est prioritaire." + "</p>" +
                                        "<p>Merci de votre compréhension." + "</p>" +
                                        // "<p>Du : " +formatDebut + "</p>" +
                                        // "<p>Au : " +FormatFin + "<p>"+

                                        "<p>Bonne journée," + "</p>";
                                          var emailTo =data[b][5]; // personnalisation du mail d'envoi 
                                          var Subject = "Annulation Réservation Kangoo ";

                                          MailApp.sendEmail({
                                                  to: emailTo,
                                                  cc: "",
                                                  subject: Subject,
                                                  htmlBody: messageConfirme,});}

              }
            }


        //-------------------------------------------------------------------------------------------------
        //-----------------------------------Boucle attribution voiture------------------------------------
        //-------------------------------------------------------------------------------------------------
        if (Mail != "" && Vehicule == "" && ServiceInfo != "Oui") {

            var checkKm = [];
            var reservedCars = [];
            var cars = ["Porche", "Vélo", "A pied","Kangoo"]; // mettre véhicule dans l'ordre voulu


            for (var a = 0; a < data.length; a++) {
                var currentRow = data[a];
                var sdate = currentRow[1];		//récupération des données de la ligne en cours
                var edate = currentRow[2];
              
                if ((DateDebut >= sdate && DateDebut <= edate) || (DateFin >= sdate && DateFin <= edate)) {
                    reservedCars.push(currentRow[8]);  
                    checkKm.push(currentRow[6]);
                }
            }


            let availableCars = cars.filter(x => reservedCars.indexOf(x) === -1);

            if(availableCars != "") { // rajout

            sheet.getRange(startcolumn + i, 9).setValue(availableCars[0]);

            Vehicule = availableCars[0];

            } // rajout --------------------------------------
            else {
              checkKm.sort(compare);
              if ( Distance > checkKm[0] && currentRow[7] != "Oui") {
                for (var c =0; c < data.length; c++) {
                  currentRow = data[c];

                  if (currentRow[6] === checkKm[0]) {
                    var cancelpersonne = currentRow[3];
                    var cancelMail = currentRow[5];
                    var cancelCar = currentRow[8];
                    sheet.getRange(startcolumn + c, 13).setValue("x");
                    sheet.getRange(startcolumn + c, 12).setValue("Oui");
                    sheet.getRange(startcolumn + c, 11).setValue("TERMINE"); //voir simplifier en une colonne (annulé et terminer)
                    sheet.getRange(startcolumn + i, 9).setValue(cancelCar);
                    Vehicule = cancelCar

                var sendTo = cancelMail;
       	        var messageConfirme = "";
                                        messageConfirme += "<p>Bonjour, " +cancelpersonne + " </p>" +
                                        "<p>Nous vous informons de l'annulation du véhicule : "+ ""+ cancelCar + "</p>" +
                                        "<p>Merci de votre compréhension." + "</p>" +
                                        // "<p>Du : " +formatDebut + "</p>" +
                                        // "<p>Au : " +FormatFin + "<p>"+

                                        "<p>Bonne journée," + "</p>";
                                          var emailTo =cancelMail; // personnalisation du mail d'envoi 
                                          var Subject = "Annulation Réservation " +cancelCar;

                                          MailApp.sendEmail({
                                                  to: emailTo,
                                                  cc: "",
                                                  subject: Subject,
                                                  htmlBody: messageConfirme,});}

                  }
                }
              }
            }

        

        //-------------------------------------------------------------------------------------------------
        //-------------------------------------Envoi Mail confirmation-------------------------------------
        //-------------------------------------------------------------------------------------------------


        var formattedDebut = Utilities.formatDate(new Date(DateDebut), "GMT +1", "dd/MM/yyyy");
        var FormattedFin = Utilities.formatDate(new Date(DateFin), "GMT + 1", "dd/MM/yyyy");

        var description = Prenom + " " + Vehicule + " " + Distance   // concatenation des champs des colonnes  E H destinée à la zone Description de l'agenda

        if (sheet.getRange, startcolumn + i, 10 != "AJOUTE" && DateDebut != "") {
            var eventImported = column[9];// Colonne I - Statut de l'importation - colonne AJOUTE A L'AGENDA

            var setupInfo = ss.getSheetByName("agenda"); // Nom de la feuille de calcul contenant la nom de l'agenda
            var calendarName = setupInfo.getRange("A1").getValue(); // Référence de la cellule contenant le nom de l'agenda


            Utilities.sleep(500);


            if (sheet.getRange(startcolumn + i, 10) == "AJOUTE") {
                break
            }


            if (eventImported != EVENT_IMPORTED && DateDebut != "") {  // Evite les doublons dans l'agenda, si le texte AJOUTE est présent en I,       l'événement   n'est pas ajouté
                var cal = CalendarApp.openByName(calendarName);


                cal.createEvent(Prenom + " " + Vehicule, new Date(DateDebut), new Date(DateFin), { description: description }); // Création de l'événement dans l'agenda avec le titre, la date de début, la date de fin et la description complète


                var currentRow = data[i]
                var sendTo = currentRow[6]
                var messageConfirme = "";
                messageConfirme += "<p>Bonjour" + "" + currentRow[3]+ "," + "</p>" +
                    "<p>Nous vous confirmons le réservation du véhicule (sous réserve d'annulation) : " + Vehicule + "</p>" +
                    "<p>Du : " + formattedDebut + "</p>" +
                    "<p>Au : " + FormattedFin + "<p>" +
                    "<p>Pour tout annulation ou changement, merci de supprimer votre réservation sur l'Agenda du Drive. Merci<p>"
                "<p>Bonne journée," + "</p>";
                var emailTo = 'johnnylombard@hotmail.fr'; // personnalisation du mail d'envoi
                var Subject = "Réservation " + " " + Prenom + " " + Vehicule;

                MailApp.sendEmail({
                    to: emailTo,
                    cc: "",
                    subject: Subject,
                    htmlBody: messageConfirme,
                });


                sheet.getRange(startcolumn + i, 10).setValue(EVENT_IMPORTED);


                SpreadsheetApp.flush();
            }
        }

    }
}
