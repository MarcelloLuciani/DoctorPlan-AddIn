
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

	// Funzione che gestisce la risoluzione del problema dei turni dei chirurghi
async function risolviClingo() {
    try {

        // 1. Lettura dati da Excel
        const mappaDati = await leggiTabella();

        console.log("Dati letti:", mappaDati); // Stampa i dati letti per il debug

        // Verifica se sono stati letti dei dati
        if (mappaDati.length === 0) {
            console.error("Errore: Nessuna tabella trovata o nessun dato letto.");
            return; // Esci dalla funzione se non ci sono dati
        }

        // 2. Eseguo Clingo
        let risposta = await eseguiClingoWasm(mappaDati); 


        // 5. Mostri i risultati
        console.log(risposta);
        mostraRisultati(risposta);

        
        // 6. Crea la tabella di risposta
        createScheduleFromClingo(risposta);
      
    } catch (error) {
        errorHandler(error);
    }
}

	// Funzione principale che genera la tabella turni in Excel
async function createScheduleFromClingo(clingoResponse) {
    try {
        console.log("Elaborazione risposta di Clingo:", clingoResponse);
        await generateScheduleTable(clingoResponse);
        console.log("Tabella dei turni generata con successo nel foglio 'Turni'!");
    } catch (error) {
        console.error("Errore nella generazione della tabella:", error);
        console.log("Tipo di dato ricevuto:", typeof clingoResponse);
        if (typeof clingoResponse === 'string') {
            console.log("Primi 100 caratteri:", clingoResponse.substring(0, 100));
        } else {
            console.log("Struttura completa:", JSON.stringify(clingoResponse));
        }
    }
}