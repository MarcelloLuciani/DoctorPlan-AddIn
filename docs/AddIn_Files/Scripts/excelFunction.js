    // Funzione per creare una tabella con chirurghi e turni - EXCEL
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

        if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
            sheet.getUsedRange().format.autofitColumns();
            sheet.getUsedRange().format.autofitRows();

        } else {

            sheet.getRange("A:A").format.columnWidth = 150;
            sheet.getRange("B:B").format.columnWidth = 150;
            sheet.getRange("C:C").format.columnWidth = 150;
            sheet.getRange("D:D").format.columnWidth = 150;
        }
        
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

	// Funzione che gestisce il cambiamento nelle celle della tabella - EXCEL
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

    // Funzione per eliminare la tabella con intestazione "Chirurghi" e "Turni" - EXCEL
function deleteTable() {
    Excel.run(async function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var tables = sheet.tables; // Ottiene tutte le tabelle nel foglio
        tables.load("items/name"); // Carica i nomi delle tabelle

        await context.sync(); // Sincronizza per ottenere i dati

        for (let table of tables.items) {
            
            if (table.name === "TabellaChirurghiTurni") {
                let tableRange = table.getRange(); // Ottiene l'intervallo della tabella
                tableRange.format.autofitColumns();
                tableRange.format.columnWidth = 48;
                await context.sync(); // Assicuriamoci che la formattazione sia applicata

                // Dopo aver effettuato l'autofit, possiamo eliminare la tabella
                table.delete(); // Elimina la tabella
                console.log("Tabella Chirurghi - Turni eliminata con successo!");
            }
        }

        return context.sync(); // Ultima sincronizzazione per garantire che tutto sia stato applicato correttamente

    }).catch(function (error) {
            console.error("Error: " + error);
        });
}

    // Funzione per la lettura dei dati da Excel - EXCEL
async function leggiTabella() {
    try {
        return await Excel.run(async (context) => {

            // Recupero il foglio di lavoro attivo
            var sheet = context.workbook.worksheets.getActiveWorksheet();
            var tables = sheet.tables; // Ottiene tutte le tabelle nel foglio
            tables.load("items/name"); // Carica i nomi delle tabelle

            await context.sync(); // Sincronizza per ottenere i dati

            for (let table of tables.items) {
                if (table.name === "TabellaChirurghiTurni") {
                    console.log("Tabella trovata:", table.name);

                    // Get data from columns
                    let columnName = table.columns.getItem("Nome Chirurgo").getDataBodyRange().load("values");
                    let columnTurn = table.columns.getItem("Disponibilità Lavorativa").getDataBodyRange().load("values");
                    let columnPref = table.columns.getItem("Preferenza di Turno").getDataBodyRange().load("values");
                    let columnDisp = table.columns.getItem("Qta Turni").getDataBodyRange().load("values");

                    // Sync to populate proxy objects with data from Excel
                    await context.sync();

                    let nameColumnValues = columnName.values;
                    let turnColumnValues = columnTurn.values;
                    let preferenceColumnValues = columnPref.values;
                    let disponibilityColumnValues = columnDisp.values;

                    // Sync to update the sheet in Excel
                    await context.sync();

                    let doctors = [];

                    for (let i = 0; i < nameColumnValues.length; i++) {

                        let doctor = {
                     
                            "name": formatter(String(nameColumnValues[i])),
                            "turn": formatter([nameColumnValues[i],turnColumnValues[i], "turn"]),
                            "preference": formatter([nameColumnValues[i], preferenceColumnValues[i], "preference"]), 
                            "disponibility": formatter([nameColumnValues[i], disponibilityColumnValues[i], "disponibility"]), 

                        }
                        doctors.push(doctor);

                    }

                    // Create an array to hold the data and return it
                    return doctors;
                }
            }

            return null; // se la tabella non viene trovata
        });
    } catch (error) {
        console.error("Errore nella lettura della tabella:", error);
        return null;
    }
}

    // Funzione per mostrare i risultati di Clingo in un nuovo foglio creato appositamente - EXCEL
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
            const nuovoFoglio = context.workbook.worksheets.add("Risultato");

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

    // Funzione per creare una tabella con la risposta ricevuta da Clingo - EXCEL
function generateScheduleTable(clingoResponse) {
    // Verifica se clingoResponse è già un oggetto o è una stringa JSON
    const data = typeof clingoResponse === 'string' ? JSON.parse(clingoResponse) : clingoResponse;

    console.log("Struttura dati ricevuta:", JSON.stringify(data, null, 2));

    // Controlliamo la struttura esatta dei dati per accedere correttamente ai valori
    let solution = [];

    // Verifica se i dati sono in formato Array con un indice [0]
    if (Array.isArray(data) && data[0] && data[0].Witnesses && data[0].Witnesses[0] && data[0].Witnesses[0].Value) {
        solution = data[0].Witnesses[0].Value;
    }
    // Verifica se i dati hanno la struttura Witnesses direttamente nell'oggetto principale
    else if (data.Witnesses && data.Witnesses[0] && data.Witnesses[0].Value) {
        solution = data.Witnesses[0].Value;
    }
    // Verifica se i dati hanno la struttura Call[0].Witnesses
    else if (data.Call && data.Call[0] && data.Call[0].Witnesses && data.Call[0].Witnesses[0] && data.Call[0].Witnesses[0].Value) {
        solution = data.Call[0].Witnesses[0].Value;
    } else {
        console.error("Formato dati non riconosciuto:", data);
        throw new Error("Formato dati Clingo non supportato");
    }

    // Definisci i giorni della settimana in ordine
    const giorni = ["lun", "mar", "mer", "giov", "ven", "sab"];

    // Estrai le assegnazioni dal formato assegna(chirurgo,turno,giorno)
    const assegnazioni = solution.map(atom => {
        const match = atom.match(/assegna\((.*?),(.*?),(.*?)\)/);
        if (match) {
            return {
                chirurgo: match[1],
                turno: match[2],
                giorno: match[3]
            };
        }
        return null;
    }).filter(item => item !== null);

    console.log("Assegnazioni estratte:", assegnazioni);

    // Crea una mappa dei chirurghi con i loro turni
    const chirurghi = {};
    assegnazioni.forEach(assegnazione => {
        if (!chirurghi[assegnazione.chirurgo]) {
            chirurghi[assegnazione.chirurgo] = {
                turno: assegnazione.turno,
                giorni: []
            };
        }

        // Aggiungi il giorno all'elenco dei giorni del chirurgo
        chirurghi[assegnazione.chirurgo].giorni.push(assegnazione.giorno);
    });

    console.log("Mappa dei chirurghi:", chirurghi);

    // Usa Office JS per creare la tabella Excel
    return Excel.run(async (context) => {
        // Cerca se esiste già un foglio chiamato "Turni"
        let sheet;
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");

        await context.sync();

        // Verifica se il foglio "Turni" esiste già
        let turniSheet = sheets.items.find(s => s.name === "Turni");

        if (turniSheet) {
            // Se esiste, lo cancelliamo e ricreiamo
            turniSheet.delete();
            sheet = sheets.add("Turni");
        } else {
            // Se non esiste, lo creiamo
            sheet = sheets.add("Turni");
        }

        // Pulisci il foglio per sicurezza
        sheet.getRange().clear();

        // Intestazioni
        sheet.getRange("A1").values = [["Chirurgo"]];
        sheet.getRange("B1").values = [["Turno"]];

        // Aggiungi i giorni come intestazioni
        const dayMap = {
            "lun": "Lunedì",
            "mar": "Martedì",
            "mer": "Mercoledì",
            "giov": "Giovedì",
            "ven": "Venerdì",
            "sab": "Sabato"
        };

        giorni.forEach((giorno, index) => {
            const cell = sheet.getCell(0, index + 2);
            cell.values = [[dayMap[giorno] || giorno]];
        });

        // Aggiungi i dati dei chirurghi
        let row = 1;
        Object.keys(chirurghi).forEach(nome => {
            const info = chirurghi[nome];

            // Nome chirurgo
            sheet.getCell(row, 0).values = [[nome]];

            // Turno
            sheet.getCell(row, 1).values = [[info.turno]];

            // Giorni di lavoro
            giorni.forEach((giorno, index) => {
                const cell = sheet.getCell(row, index + 2);
                if (info.giorni.includes(giorno)) {
                    cell.values = [["✓"]];
                    // Colora la cella in base al turno
                    if (info.turno === "mattina") {
                        cell.format.fill.color = "#FFEB9C"; // Giallo chiaro per mattina
                    } else if (info.turno === "pomeriggio") {
                        cell.format.fill.color = "#C5E0B4"; // Verde chiaro per pomeriggio
                    }
                } else {
                    cell.values = [[""]];
                }
            });

            row++;
        });

        // Formattazione tabella
        const tableRange = sheet.getRange("A1:H" + (row));
        tableRange.format.autofitColumns();
        tableRange.format.autofitRows();

        // Formattazione intestazioni
        const headerRow = sheet.getRange("A1:H1");
        headerRow.format.font.bold = true;
        headerRow.format.fill.color = "#4472C4";
        headerRow.format.font.color = "white";

        // Aggiunge bordi alla tabella
        tableRange.format.borders.getItem('EdgeTop').style = 'Continuous';
        tableRange.format.borders.getItem('EdgeBottom').style = 'Continuous';
        tableRange.format.borders.getItem('EdgeLeft').style = 'Continuous';
        tableRange.format.borders.getItem('EdgeRight').style = 'Continuous';
        tableRange.format.borders.getItem('InsideHorizontal').style = 'Continuous';
        tableRange.format.borders.getItem('InsideVertical').style = 'Continuous';

        // Aggiungi legenda per i colori
        sheet.getCell(row + 2, 0).values = [["Legenda:"]];
        sheet.getCell(row + 2, 0).format.font.bold = true;

        sheet.getCell(row + 3, 0).values = [["Mattina"]];
        sheet.getCell(row + 3, 1).values = [[""]];
        sheet.getCell(row + 3, 1).format.fill.color = "#FFEB9C";

        sheet.getCell(row + 4, 0).values = [["Pomeriggio"]];
        sheet.getCell(row + 4, 1).values = [[""]];
        sheet.getCell(row + 4, 1).format.fill.color = "#C5E0B4";

        // Attiva il foglio per renderlo visibile
        sheet.activate();

        await context.sync();
    });
}
