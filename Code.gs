function currentDateGS(){
   return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function sendMail(emailSubject,emailBody,receipt){
  var aubayLogoUrl  = "https://www.aubay.pt/img/aubay_smile_logo.png";
  var aubayLogoBlob = UrlFetchApp.fetch(aubayLogoUrl).getBlob().setName("aubayLogoBlob");
  
  MailApp.sendEmail({
    name: "#academia-aubay [Automatic Message]",
    to: receipt,
    subject: emailSubject,
    htmlBody: emailBody,
    inlineImages:
    {
      aubayLogo: aubayLogoBlob
    }
  });
}

let nucleos = [];
function getNucleosGS(){
    var sheetNucleos = SpreadsheetApp.openById("@PutHereTheGoogleSheetID@").getSheetByName("nucelos");
    var dataNucleos = sheetNucleos.getDataRange().getValues();
    for (var i = 1; i < dataNucleos.length; i++) {
        nucleos.push([
                        dataNucleos[i][0], // Nome
                        dataNucleos[i][1], // Email
                        dataNucleos[i][2], // Canal Slack
                        dataNucleos[i][3], // Link do Grupo Whatsapp
                        dataNucleos[i][4], // Mensagem do Welcome
                        dataNucleos[i][5], // Mensagem do Welcome Header
                        dataNucleos[i][6], // Mensagem do Welcome Footer
                        dataNucleos[i][7]  // EMail Dinamizadores
        ]);
    }
}

let dinamizadores = [];
function getDinamizadoresGS(){
    var sheetDinamizadores = SpreadsheetApp.openById("@PutHereTheGoogleSheetID@").getSheetByName("dinamizadores");
    var dataDinamizadores = sheetDinamizadores.getDataRange().getValues();
    for (var i = 1; i < dataDinamizadores.length; i++) {
        dinamizadores.push([
                        dataDinamizadores[i][0], // Nucleo
                        dataDinamizadores[i][1], // Email
                        dataDinamizadores[i][2]  // Nome
        ]);
    }
}

let colaboradores = [];
let colaboradoresQueSairam = [];
let colaboradoresQueEntraram = [];

function getColaboradoresGS(){
    getNucleosGS();
    getDinamizadoresGS();
    Utilities.sleep(100);

    var sheetColaboradores = SpreadsheetApp.openById("@PutHereTheGoogleSheetID@").getSheetByName("colaboradores");
    var dataColaboradores = sheetColaboradores.getDataRange().getValues();

    for (var i = 1; i < dataColaboradores.length; i++) {

      switch (dataColaboradores[i][0]){
        case "Novo":
          // Definindo os parâmetros para o envio do email para o colaborador
          var emailSubject = "Bem vindo à família Aubay - Nucleo " + dataColaboradores[i][8];
          var emailBody = "";
              nucleos.forEach(function(nucleo) {
                    if(nucleo[0] === dataColaboradores[i][8]){
                      emailBody = nucleo[4];
                    }
                });
          var receipt = dataColaboradores[i][10];
          
          // Enviando o email para o novo colaborador
          sendMail(emailSubject,emailBody,receipt);

          // Alterando a planilha com o status de [Novo] para [Convidado]
          var cellStatusColaborador = sheetColaboradores.getRange(i+1, 1, 1, 1);
          cellStatusColaborador.setValue('Convidado');

          // Informando a [Data do Convite] na planilha
          var cellDataConvite = sheetColaboradores.getRange(i+1, 2, 1, 1);
              cellDataConvite.setValue(currentDateGS());

          // Adicionando o colaborador no array de Novos Colaboradores; Que será enviado para o email dos dinamizadores deste nucleo
          var colaborador = '{ "status": "Novo", "data":"' + Utilities.formatDate(dataColaboradores[i][6], Session.getScriptTimeZone(), "yyyy-MM-dd") + '", "nome": "' + dataColaboradores[i][9] + '", "email": "' + dataColaboradores[i][10] + '","nucleo":"' + dataColaboradores[i][8] +'"}';
          colaboradoresQueEntraram.push(colaborador);
        break;

        case "Saiu":
          // Verificando se o campo [Data Saida] está preenchido na planiha
          if(dataColaboradores[i][2] === ""){
            // Alterando a planilha com o status de [Convidado] para [Saiu]
            var cellStatusColaborador = sheetColaboradores.getRange(i+1, 3, 1, 1);
            cellStatusColaborador.setValue(currentDateGS());

            // Adicionando o colaborador no array de Colaboradores que Saíram; Que será enviado para o email dos dinamizadores deste nucleo
          var colaborador = '{ "status": "Saiu", "data":"' + Utilities.formatDate(dataColaboradores[i][7], Session.getScriptTimeZone(), "yyyy-MM-dd") + '", "nome": "' + dataColaboradores[i][9] + '", "email": "' + dataColaboradores[i][10] + '","nucleo":"' + dataColaboradores[i][8] +'"}';
            colaboradoresQueSairam.push(colaborador);
          }
        break;                
        default:
          Logger.log("Nada a fazer");
      }
    }

    // Enviando o email para os dinamizadores dos nucleos informando-os sobre os entraram e os que saíram
    colaboradores = colaboradoresQueEntraram.concat(colaboradoresQueSairam);
    var colaboradoresJSON = colaboradores.map(JSON.parse);

    // Separando o email de acordo com cada nucleo
    [...new Set(colaboradoresJSON.map(({nucleo})=>nucleo))].forEach(function(nucleoJSON) {
        var emailDinamizadores = "";
            nucleos.forEach(function(nucleo) {
                  if(nucleo[0] === nucleoJSON){
                    emailDinamizadores = nucleo[7];
                  }
              });
  
        // Filtrando o array que possui os Colaboradores que entraram e os que sairam, pelo nucleo
        var colaboradoresPorNucleo = colaboradoresJSON.filter(function (colaboradores) {
            return colaboradores.nucleo == nucleoJSON;
        });

        // Definindo os parâmetros para o envio do email para os dinamizadores dos nucleos, informando os colaboradores que entraram e os que saíram
        var emailSubject = "Entradas e Saídas de Colaboradores - Nucleo " + nucleoJSON ;
        var emailBody = "<p>Olá Dinamizadores do Nucleo <b>" + nucleoJSON + "</b></p>" +
                        "<p>Segue lista de colaboradores que entraram e saíram:</p>";
            colaboradoresPorNucleo.forEach(function(colaboradorPorNucleo) {
                emailBody += JSON.stringify(colaboradorPorNucleo) + "<br/>";
            });
            emailBody += "<br/><p>Um forte abraço.</p>";

        var receipt = emailDinamizadores;

        // Enviando o email para os dinamizadores dos nucleos
        sendMail(emailSubject,emailBody,receipt);
    });
}
