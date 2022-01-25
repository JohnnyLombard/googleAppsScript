var EVENT_IMPORTED = "AJOUTE"; // Ajoutera le texte "AJOUTE" dans la colonne J
var ss = SpreadsheetApp.getActiveSpreadsheet();
var ActiveSheet = SpreadsheetApp.openById("1LJp4Ecau5UCK_aCA9vZiofR1__dLLl66P-6OjnPjIYM").getSheetByName("Réponses au formulaire 1");

function onOpen() {
   var menuEntries = [{name: "Ajouter les événements à l'agenda", functionName: "importCalendar"}];
   ss.addMenu("Agenda", menuEntries); // Pour ajouter une menu Agenda et un sous-menu "ajouter les événements" dans la feuille de calcul. Cela permettra de tester manuellement la liaison entre la feuille de calcul et l'agenda
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

  if(row > 1 && column === 12 && sheetName === "Réponses au formulaire 1")
  {
    if(dataSheet.getRange(row,13).getValue() === "x")
    {
      dataSheet.getRange(row,13).setValue('');
    }
    else
    {
      dataSheet.getRange(row,13).setValue('x');
    }
  }

}

      function copyRows()
{ 

  var copySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Archive");
  var dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Réponses au formulaire 1");
  var dataRange = dataSheet.getDataRange();
  var dataValues = dataRange.getValues();

  for(var i = 1; i < dataValues.length; i++)
  {
    if(dataValues[i][12] === "x")
    {
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

  for(var i = 1; i < dataValues.length; i++)
  {
    if(dataValues[i][12] === "x")
    {
      var clearRow = i+1;
      dataSheet.getRange('A' + clearRow + ':M' + clearRow).clear();
    }
  }
}

function addMenu()
{
  var menu = SpreadsheetApp.getUi().createMenu('Custom');
  menu.addItem('Copy Rows', 'copyRows');
  menu.addToUi(); 
}

// function onOpen(e)
// {
//   addMenu(); 
// }

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
          var currentRow = data[i];
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
//---------------------------réservation voiture Service INFO-------------------------------------OK
//------------------------------------------------------------------------------------------------

            if(ServiceInfo ==="Oui" && Mail !="" && Vehicule =="") {
            sheet.getRange(startcolumn + i, 9).setValue("Kangoo");}


          
//-------------------------------------------------------------------------------------------------            
//-----------------------------------Boucle attribution voiture------------------------------------
//-------------------------------------------------------------------------------------------------
            if (Mail != "" && Vehicule =="") {

            for (var a = 0; a < data.length; a++){
                var currentRow = data[a];
	            var sdate = currentRow[1];		//récupération des données
	            var edate = currentRow[2];
	            // var personne = currentRow[3];
                // var mailAnnul = currentRow[5];
	            //  var dist = currentRow[6];
	            // var info = currentRow[7];
                // var resa = currentRow[8];

                if(DateDebut < currentRow[1] && DateFin < currentRow[1] || DateDebut > currentRow[2]) {
                  sheet.getRange(startcolumn + i, 9).setValue("Véhicule 1");
            }
            }
            }




//-------------------------------------------------------------------------------------------------            
//-------------------------------------Envoi Mail confirmation-------------------------------------
//-------------------------------------------------------------------------------------------------


      var formattedDebut = Utilities.formatDate(new Date(DateDebut), "GMT +1", "dd/MM/yyyy");
      var FormattedFin = Utilities.formatDate(new Date(DateFin), "GMT + 1", "dd/MM/yyyy");

      var description =  Prenom +" "+  Vehicule +" " + Distance   // concatenation des champs des colonnes  E H destinée à la zone Description de l'agenda

      if( sheet.getRange, startcolumn + i, 10 != "AJOUTE" && DateDebut != ""){
      var eventImported = column[9];// Colonne I - Statut de l'importation - colonne AJOUTE A L'AGENDA
    
      var setupInfo = ss.getSheetByName("agenda"); // Nom de la feuille de calcul contenant la nom de l'agenda
      var calendarName = setupInfo.getRange("A1").getValue(); // Référence de la cellule contenant le nom de l'agenda


          Utilities.sleep(500);


          if (sheet.getRange(startcolumn + i, 10 ) == "AJOUTE") {
          break
          }


          if (eventImported  != EVENT_IMPORTED && DateDebut != "") {  // Evite les doublons dans l'agenda, si le texte AJOUTE est présent en I,       l'événement   n'est pas ajouté
          var cal = CalendarApp.openByName(calendarName);
          
      
          cal.createEvent(Prenom + " " + Vehicule, new Date(DateDebut), new Date(DateFin), {description : description}); // Création de l'événement dans l'agenda avec le titre, la date de début, la date de fin et la description complète
   
      

      
      var currentRow = data[i]
      var sendTo = currentRow[6]
      var messageConfirme = "";
                                              messageConfirme += "<p>Bonjour," + "</p>" +
                                              "<p>Nous vous confirmons le réservation du véhicule (sous réserve d'annulation) : "+Vehicule + "</p>" +
                                              "<p>Du : " +formattedDebut + "</p>" +
                                              "<p>Au : " +FormattedFin + "<p>"+
                                              "<p>Pour tout annulation ou changement, merci de supprimer votre réservation sur l'Agenda du Drive. Merci<p>"
                                              "<p>Bonne journée," + "</p>";
                                          var emailTo = 'johnnylombard@hotmail.fr'; // personnalisation du mail d'envoi 
                                          var Subject = "Réservation "+ " " +Prenom+ " " + Vehicule;
        
                                          MailApp.sendEmail({
                                                  to: emailTo,
                                                  cc: "",
                                                  subject: Subject,
                                                  htmlBody: messageConfirme,});


                sheet.getRange(startcolumn + i, 10).setValue(EVENT_IMPORTED);
                                      

      SpreadsheetApp.flush();  
    }}
  
}
}

//  if(ServiceInfo ==="Oui" && Mail !="" && Vehicule =="") {
//             sheet.getRange(startcolumn + i, 9).setValue("Kangoo");}

//             if(Mail !="" && Vehicule =="" && ServiceInfo !="Oui") {
              
              
//                 sheet.getRange(startcolumn + i, 9).setValue("Véhicule 1");
//               } 
            









//           for (var j = 0; j < data.length; j++) {
//             var currentRow = data[j];
// 	          var sdate = currentRow[1];		//récupération des données
// 	          var edate = currentRow[2];
// 	          var personne = currentRow[3];
//             var mailAnnul = currentRow[5];
// 	          var dist = currentRow[6];
// 	          var info = currentRow[7];
//             var resa = currentRow[8];
          

//             if(Vehicule == "Kangoo" && resa == "Kangoo" && sdate <= DateFin <= edate && info == "Non" && data[i][12] !="x" && data[j][8] != "") {


//             sheet.getRange(startcolumn + i, 13).setValue("x");
//             sheet.getRange(startcolumn + i, 12).setValue("oui");
//             sheet.getRange(startcolumn + i, 11).setValue("TERMINE");
//             sheet.getRange(startcolumn + i, 10).setValue(EVENT_IMPORTED);
//               var sendTo = currentRow[5];
//        	      var messageConfirme = "";
//                                         messageConfirme += "<p>Bonjour, " +currentRow[4] + " </p>" +
//                                         "<p>Nous vous informons de l'annulation de votre réservation pour le véhicule : "+resa + ", car le service Informatique est prioritaire." + "</p>" +
//                                         "<p>Merci de votre compréhension." + "</p>" +
//                                         // "<p>Du : " +formatDebut + "</p>" +
//                                         // "<p>Au : " +FormatFin + "<p>"+
                                              
//                                         "<p>Bonne journée," + "</p>";
//                                           var emailTo = mailAnnul; // personnalisation du mail d'envoi 
//                                           var Subject = "Annulation Réservation " + resa;
        
//                                           MailApp.sendEmail({
//                                                   to: emailTo,
//                                                   cc: "",
//                                                   subject: Subject,
//                                                   htmlBody: messageConfirme,});}

//             }