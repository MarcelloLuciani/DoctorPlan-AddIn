
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

			applyOfficeTheme();
            

            $('#bottone-func-1').on('click', helloWorld);
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

// Funzione writeHelloWorld 
function helloWorld() {
    Excel.run(function (context) {
        // Ottieni la cella A1
        var range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");

        // Imposta il valore "Hello World"
        range.values = [["Hello World"]];

        // Applica le modifiche
        return context.sync();

    }).catch(function (error) {
        errorHandler(error);
    });
}

// Funzione per creare una tabella con chirurghi e turni
function createSurgeonShiftTable(numRows) {
    Excel.run(function (context) {
        // Recuperiamo il foglio di lavoro attivo
        var sheet = context.workbook.worksheets.getActiveWorksheet();

        // Definiamo il range per la tabella (iniziamo dalla cella A1)
        var headerRange = sheet.getRange("A1:B1");

        // Impostiamo le intestazioni della tabella
        headerRange.values = [["Chirurghi", "Turni"]];

        // Formattazione delle intestazioni
        headerRange.format.font.bold = true;
        headerRange.format.fill.color = "#4472C4";  // Colore blu
        headerRange.format.font.color = "white";

        // Calcoliamo il range della tabella completa
        var fullRangeAddress = "A1:B" + (numRows + 1);
        var fullRange = sheet.getRange(fullRangeAddress);

        // Crea una tabella con le intestazioni
        var table = sheet.tables.add(fullRange, true);
        table.name = "TabellaChirurghiTurni";

        // Sistemo la larghezza delle colonne per visualizzare meglio i dati
        sheet.getRange("A:A").format.columnWidth = 150;
        sheet.getRange("B:B").format.columnWidth = 150;

        // Creiamo un array di righe vuote per la tabella
        var data = [];
        for (var i = 0; i < numRows; i++) {
            data.push(["", ""]);
        }

        // Aggiungiamo le righe alla tabella
        if (numRows > 0) {
            var dataRange = sheet.getRange("A2:B" + (numRows + 1));
            dataRange.values = data;
        }

        // Nella colonna "Turni" voglio obbligare l'utente a scegliere tra le opzioni suggerite
        var shiftRange = sheet.getRange("B2:B" + (numRows + 1));
        var validation = shiftRange.dataValidation;
        validation.rule = {
            list: {
                inCellDropDown: true,
                source: "mattina,pomeriggio,sera"
            }
        };

        return context.sync();

    }).catch(function (error) {
        errorHandler(error)
    });
}

// Funzione per eliminare la tabella con intestazione "Chirurghi" e "Turni"
function deleteTable() {
    Excel.run(async function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var tables = sheet.tables; // Ottiene tutte le tabelle nel foglio
        tables.load("items/name"); // Carica i nomi delle tabelle

        await context.sync(); // Sincronizza per ottenere i dati

        for (let table of tables.items) {
            table.columns.load("items/name"); // Carica i nomi delle colonne
            await context.sync(); // Sincronizza prima di accedere ai dati

            let columnNames = table.columns.items.map(col => col.name);

            // Controlla se l'intestazione è esattamente ["Chirurghi", "Turni"]
            if (JSON.stringify(columnNames) === JSON.stringify(["Chirurghi", "Turni"])) {
                let tableRange = table.getRange(); // Ottiene l'intervallo della tabella
                tableRange.format.autofitColumns();
                tableRange.format.columnWidth = 48;
                await context.sync(); // Assicuriamoci che la formattazione sia applicata

                // Dopo aver effettuato l'autofit, possiamo eliminare la tabella
                table.delete(); // Elimina la tabella
                console.log("Tabella eliminata con intestazione Chirurghi - Turni");
            }
        }

        return context.sync(); // Ultima sincronizzazione per garantire che tutto sia stato applicato correttamente
    }).catch(function (error) {
        console.error("Error: " + error);
    });
}

/* Test di Clingo per verificarne l'esatto funzionamento
async function runClingo() {
    try {
        // Verifica che 'clingo' sia definito
        console.log("Verifica se 'clingo' è definito:", typeof clingo);

        // Verifica se clingo è definito
        if (typeof clingo === 'undefined') {
            throw new Error("clingo is not defined. Please ensure that clingo-wasm is loaded.");
        }

        // Inizializza Clingo con il file WASM (opzionale, se necessario)
        await clingo.init("https://cdn.jsdelivr.net/npm/clingo-wasm@0.2.1/dist/clingo.wasm");

        // Esegui un programma di esempio su Clingo
        const result1 = await clingo.run("a. b :- a.");
        console.log("Risultato 1:", result1);

        const result2 = await clingo.run("{a; b; c}.", 0);
        console.log("Risultato 2:", result2);

        // Mostra i risultati nel div con id "results"
        $('#results').html(`
            <p><strong>Risultato 1:</strong> ${JSON.stringify(result1)}</p>
            <p><strong>Risultato 2:</strong> ${JSON.stringify(result2)}</p>
        `);

    } catch (error) {
        console.error("Errore durante l'esecuzione di Clingo:", error);
        $('#results').html('<p>Si è verificato un errore durante l\'esecuzione di Clingo.</p>');
    }
}
*/

// Funzione che gestisce la risoluzione del problema dei turni dei chirurghi
async function risolviClingo() {
    try {
        // 1. Lettura dati da Excel
        const datiTurni = await leggiDatiTurni();

        // Verifica se sono stati letti dei dati
        if (!datiTurni || datiTurni.length === 0) {
            console.error("Errore: Nessuna tabella trovata o nessun dato letto.");
            return; // Esci dalla funzione se non ci sono dati
        }

        // 2. Creazione File con dati da passare a clingo
        await scriviDatiToNuovoFoglio(datiTurni);

        // 3. Leggo i dati formattati
        let dati = await leggiDatiDalFoglio();

        // Verifica se i dati letti sono validi
        if (!dati || dati.length === 0) {
            console.error("Errore: Nessun dato formattato trovato nel foglio.");
            return; // Esci dalla funzione se non ci sono dati
        }

        // 4. Eseguo Clingo
        let risposta = await eseguiClingoWasm(dati); 

        // 5. Mostri i risultati
        console.log(risposta);
        mostraRisultati(risposta);

    } catch (error) {
        errorHandler(error);
    }
}

// Funzione per la lettura dei dati da Excel
async function leggiDatiTurni() {
    return Excel.run(async (context) => {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var tables = sheet.tables; // Ottiene tutte le tabelle nel foglio
        tables.load("items/name"); // Carica i nomi delle tabelle

        await context.sync(); // Sincronizza per ottenere i dati

        let datiTurni = []; // Array per raccogliere i dati

        // Scorri tutte le tabelle
        for (let table of tables.items) {
            table.columns.load("items/name"); // Carica i nomi delle colonne
            await context.sync(); // Sincronizza prima di accedere ai dati

            let columnNames = table.columns.items.map(col => col.name);

            // Controlla se l'intestazione è esattamente ["Chirurghi", "Turni"]
            if (JSON.stringify(columnNames) === JSON.stringify(["Chirurghi", "Turni"])) {
                const range = table.getDataBodyRange(); // Ottieni i dati
                range.load("values");
                await context.sync();

                // Aggiungi i valori letti nella variabile datiTurni
                datiTurni = range.values;
                break; // Interrompi il ciclo dopo aver trovato la tabella corretta
            }
        }

        // Restituisci i dati letti (se trovati)
        return datiTurni;
    });
}

// Funzione per scrivere i dati in un nuovo foglio creato appositamente
async function scriviDatiToNuovoFoglio(dati) {
    return Excel.run(async (context) => {
        // Aggiungi un nuovo foglio di lavoro
        const nuovoFoglio = context.workbook.worksheets.add("Risultati Formattati");

        // Crea l'intervallo dinamicamente in base al numero di righe
        const intervalloRisultati = nuovoFoglio.getRange("A1").getResizedRange(dati.length - 1, 0);

        // Formatta i dati nel formato "(nome chirurgo, turno)"
        const datiFormattati = dati.map(row => {
            return [`(${row[0]}, ${row[1]}).`]; // Formatta ogni riga come "(nome chirurgo, turno)"
        });

        // Scrivi i dati formattati nel nuovo foglio
        intervalloRisultati.values = datiFormattati;

        await context.sync(); // Sincronizza per applicare le modifiche
        console.log("Dati formattati e scritti nel nuovo foglio.");
    });
}

// Funzione per leggere i dati dal foglio (potrebbe essere evitata caricando direttamente i dati nella funzione precedente)
async function leggiDatiDalFoglio() {
    return Excel.run(async (context) => {
        // Recupero il foglio "Risultati Formattati"
        const sheet = context.workbook.worksheets.getItem("Risultati Formattati");

        let riga = 1; 
        let datiLetti = [];

        while (true) {
            
            const range = sheet.getRange(`A${riga}`);
            range.load("values");

            await context.sync(); 

            // Se la cella è vuota, esco dal ciclo
            if (range.values[0][0] === "" || range.values[0][0] === null) {
                break; 
            }

            // Aggiungi il valore della cella all'array dei dati
            datiLetti.push(range.values[0][0]);

            // Incrementa la riga per leggere la successiva
            riga++;
        }
        return datiLetti;
    });
}

// Funzione per eseguire Clingo con i dati letti dal foglio
async function eseguiClingoWasm(datiFormattati) {
    try {
        // Inizializza Clingo con il file WASM 
        await clingo.init("https://cdn.jsdelivr.net/npm/clingo-wasm@0.2.1/dist/clingo.wasm");

        // Preparo il programma di Clingo di base
        const scriptClingo = `
giorno(lun).
giorno(mar).
giorno(mer).
giorno(giov).
giorno(ven).
giorno(sab).

1={orario(Chirurgo,G):chirurgo(Chirurgo,_)} :- giorno(G).
        `;

         // Modifico i dati letti dal foglio in regole valide per Clingo
         const datiClingo = datiFormattati.map((entry) => {

             
            const match = entry.match(/^\(([^,]+),\s*([^,]+)\)\.$/);

             if (match) {

                const chirurgo = match[1].trim().replace(/['"]+/g, ''); 
                const turno = match[2].trim().replace(/['"]+/g, ''); 

                // Faccio un controllo sui dati letti per evitare dati vuoti
                if (chirurgo && turno) {
                    return `
chirurgo('${chirurgo}', ${turno}).
                `;
                }
            }
             return ''; // In caso di dati non validi, restituisco una stringa vuota
        }).join("\n");

        // Creo il programma completo di Clingo
        const programmaCompleto = scriptClingo + "\n" + datiClingo;

        console.log("Programma completo di Clingo:", programmaCompleto);

        // Eseguo lo script
        const risultato = await clingo.run(programmaCompleto);
        console.log("Risultato di Clingo:", risultato);

        return risultato;
    } catch (error) {
        errorHandler(error);
    }
}

// Funzione per mostrare i risultati di Clingo in un nuovo foglio creato appositamente
async function mostraRisultati(risultatoClingo) {
    try {
        await Excel.run(async (context) => {
            // Recupero tutti i fogli di lavoro ed elimino quello "Temporaneo" creato per i dati
            const fogli = context.workbook.worksheets;
            const foglioEsistente = fogli.getItemOrNullObject("Risultati Formattati");
            await context.sync();

            if (!foglioEsistente.isNullObject) {
                foglioEsistente.delete();
            }

            //Creo un nuovo foglio per i risultati
            const nuovoFoglio = context.workbook.worksheets.add("Orario");

            // Converto i risultati in stringhe leggibili
            const output1 = JSON.stringify(risultatoClingo);

            
            // Aggiungo un' intestazione
            const intestazione = [["Risultati"]];

            // Scrivo i risultati
            nuovoFoglio.getRange("A1").values = intestazione;
            const cella = nuovoFoglio.getRange("A2");
            cella.values = [[output1]];
            cella.format.wrapText = true;
            cella.format.columnWidth = 350;
            cella.format.rowHeight = 150;
            

            await context.sync();
            console.log("Nuovo foglio con i risultati di Clingo creato.");
        });
    } catch (error) {
        errorHandler(error);
    }
}

// Funzione per recuperare il tema dell'applicazione e adattare il css di conseguenza
function applyOfficeTheme() {
    // Identify the current Office theme in use.
    const currentOfficeTheme = Office.context.officeTheme.themeId;

    if (currentOfficeTheme === Office.ThemeId.Colorful || currentOfficeTheme === Office.ThemeId.White) {
        console.log("No changes required.");
    }

    // Get the colors of the current Office theme.
    const bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
    const bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
    const controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor;
    const controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

    // Apply theme colors to a CSS class.
    $("body").css("background-color", bodyBackgroundColor);

    // Imposta il colore del testo in base alla luminosità del tema
    if (Office.context.officeTheme.isDarkTheme && Office.context.officeTheme.isDarkTheme()) {
        $("h1, p, span").css("color", "white");
    } else {
        $("h1, p, span").css("color", "black");
    }
}
}