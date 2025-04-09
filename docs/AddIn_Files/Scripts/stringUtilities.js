// Funzione per formattare i dati letti da Excel in regole valide per Clingo
function formatter(input) {

    let datiFormattati;

    /*
        Risultato atteso:

        chirurgo(mario). // 1.Colonna della tabella
        disponibilita(mario, mattina). // 2.Colonna della tabella
        preferenza(mario, pomeriggio). // 3.Colonna della tabella
        quantita(mario, 2). // 4.Colonna della tabella


        Logica:
            - se un chirurgo ha dichiarato in disponibilità ENTRAMBE allora devono essere generate due uscite:
                disponibilita(mario, mattina).
                disponibilita(mario, pomeriggio).

            - se un chirurgo ha dichiarato in preferenza NESSUNA allora non deve essere generate alcuna uscita
    */

    if (typeof input === 'string') {

        datiFormattati = "(" + String(input.replace(/[/[/]']/g, "")) + ").";

    } else if (Array.isArray(input)) {

        let nome = String(input[0]).replace(/[/[/]']/g, "");
        let componente = String(input[1]).replace(/[/[/]']/g, "");


        if (componente === "Entrambe") { // Caso 1 : Disponibilità

            datiFormattati = "(" + nome + ", mattina)." + " (" + nome + ", pomeriggio).";

        } else if (componente !== "Nessuna") {  // Caso 2 : Preferenza

            datiFormattati = "(" + nome + ", " + componente + ").";


        } else { // componente == "Nessuna" 

            datiFormattati = "";
        }



    } else
        console.error("Tipo di dato non supportato:", typeof input);

    return datiFormattati;

}
