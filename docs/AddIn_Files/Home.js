let messageBanner;


    // Initialization when Office JS and JQuery are ready.
    Office.onReady(() => {
        $(() => {
            // Initialize he Office Fabric UI notification mechanism and hide it.
            const element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();
            
            //// If not using Excel 2016 or later, use fallback logic.
            //if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
            //    $("#template-description").text("This sample will display the value of the cells that you have selected in the spreadsheet.");
            //    $('#button-text').text("Display!");
            //    $('#button-desc').text("Display the selection");

            //    $('#highlight-button').on('click',displaySelectedCells);
            //    return;
            //}


            // Imposta il testo per il pulsante Hello World
            //$('#button-text1').text("Hello World!");
            //$('#button-desc1').text("Writes Hello World in cell A1");



            //Funzioni per gestire la tabella
            $('#cancel-table-button').on('click', deleteTable);

            // Gestore per il pulsante di conferma nel form
            $('#confirm-table-button').on('click', function () {
                // Ottieni il numero di righe dal campo di input
                const rows = parseInt($('#table-rows').val()) || 2; // Default a 2 se non valido

                // Crea la tabella con il numero di righe specificato
                createSurgeonShiftTable(rows);
            });


            $('#btnRisolvi').on('click', risolviClingo);
            $('#btnNewFuncTest').on('click', newFunctTest);

        });
    });

	// Helper function for treating errors
	function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
	}

	// Helper function for displaying notifications
	function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
	}
	
	// Funzione per creare una tabella con chirurghi e turni
	function createSurgeonShiftTable(numRows) {
		Excel.run(function (context) {
        // Recuperiamo il foglio di lavoro attivo
        var sheet = context.workbook.worksheets.getActiveWorksheet();

        // Definiamo il range per la tabella (iniziamo dalla cella A1)
        var headerRange = sheet.getRange("A1:D1");

        // Impostiamo le intestazioni della tabella
        headerRange.values = [["Nome Chirurgo", "Disponibilità Lavorativa", "Preferenza di Turno", "Qta Turni"]];

        // Formattazione delle intestazioni
        headerRange.format.font.bold = true;
        headerRange.format.fill.color = "#4472C4";  // Colore blu
        headerRange.format.font.color = "white";

        // Calcoliamo il range della tabella completa
        var fullRangeAddress = "A1:D" + (numRows + 1);  // Corretto a D (4 colonne)
        var fullRange = sheet.getRange(fullRangeAddress);

        // Crea una tabella con le intestazioni
        var table = sheet.tables.add(fullRange, true);
        table.name = "TabellaChirurghiTurni";

        // Sistemo la larghezza delle colonne per visualizzare meglio i dati
        sheet.getRange("A:A").format.columnWidth = 150;
        sheet.getRange("B:B").format.columnWidth = 150;
        sheet.getRange("C:C").format.columnWidth = 150;
        sheet.getRange("D:D").format.columnWidth = 150;

        // Creiamo un array di righe vuote per la tabella
        var data = [];
        for (var i = 0; i < numRows; i++) {
            data.push(["", "", "", ""]);  // Quattro valori vuoti per ogni riga
        }

        // Aggiungiamo le righe alla tabella
        if (numRows > 0) {
            var dataRange = sheet.getRange("A2:D" + (numRows + 1));
            dataRange.values = data;
        }

        // Nella colonna "Turni Disponibili" voglio obbligare l'utente a scegliere tra le opzioni suggerite
        var turniDisponibiliRange = sheet.getRange("B2:B" + (numRows + 1));
        var turniDisponibiliValidation = turniDisponibiliRange.dataValidation;
        turniDisponibiliValidation.rule = {
            list: {
                inCellDropDown: true,
                source: "Mattina,Pomeriggio,Entrambe"
            }
        };

        // Nella colonna "Preferenza di Turno" voglio obbligare l'utente a scegliere tra le opzioni suggerite
        var preferenzaTurnoRange = sheet.getRange("C2:C" + (numRows + 1));  // Corretto a C
        var preferenzaTurnoValidation = preferenzaTurnoRange.dataValidation;
        preferenzaTurnoValidation.rule = {
            list: {
                inCellDropDown: true,
                source: "Mattina,Pomeriggio,Nessuna"
            }
        };

        // Aggiunta validazione per "Qta Turni" (numeri interi tra 1 e 7)
        var qtaTurniRange = sheet.getRange("D2:D" + (numRows + 1));
        var qtaTurniValidation = qtaTurniRange.dataValidation;
        qtaTurniValidation.rule = {
            wholeNumber: {
                formula1: 1,
                formula2: 7,
                operator: Excel.DataValidationOperator.between
            }
        };

        // Aggiunta dell'evento per sincronizzare la colonna B con la colonna C
        sheet.onChanged.add(handleTableChange);

        return context.sync();
    }).catch(function (error) {
        errorHandler(error);
    });
}

	// Funzione che gestisce il cambiamento nelle celle della tabella
	function handleTableChange(event) {
		return Excel.run(function (context) {
        // Verifichiamo se il cambiamento è avvenuto nella colonna B (Turni Disponibili)
        var changedRange = context.workbook.worksheets.getActiveWorksheet().getRange(event.address);
        changedRange.load(["columnIndex", "values", "columnCount", "rowIndex", "rowCount"]);

        return context.sync().then(function () {
            // Controlliamo se la modifica è nella colonna B (indice 1)
            if (changedRange.columnIndex === 1 && changedRange.columnCount === 1) {
                var valore = changedRange.values[0][0];
                var riga = changedRange.rowIndex;

                // Aggiorniamo la colonna C nella stessa riga solo se il valore è Mattina o Pomeriggio
                if (valore === "Mattina" || valore === "Pomeriggio") {
                    var preferenceCell = context.workbook.worksheets.getActiveWorksheet().getCell(riga, 2); // Colonna C (indice 2)
                    preferenceCell.values = [[valore]]; // Impostiamo lo stesso valore

                    return context.sync();
                } else if (valore === "Entrambe") {
                    // Se viene selezionato "Entrambe", non modifichiamo automaticamente la preferenza
                    // L'utente dovrà scegliere manualmente
                    return context.sync();
                }
            }
        });
    }).catch(function (error) {
        console.error("Errore: " + error);
    });
}