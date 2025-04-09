    // Funzione per formattare i dati letti da Excel in regole valide per Clingo - STRING
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

        datiFormattati = "chirurgo(" + String(input.replace(/[/[/]']/g, "")).toLowerCase() + ").";

    } else if (Array.isArray(input)) {

        let nome = String(input[0]).replace(/[/[/]']/g, "").toLowerCase();
        let componente = String(input[1]).replace(/[/[/]']/g, "").toLowerCase();


        if (input[2] === "turn") {  //Caso 1: Turno

            if (componente === "entrambe") {
                datiFormattati = "turno(" + nome + ", mattina). " +
                    "turno(" + nome + ", pomeriggio).";
            } else {
                datiFormattati = "turno(" + nome + ", " + componente + ").";
            }

        } else if (input[2] == "preference") { //Caso 2: Preferenza

            if (componente !== "nessuna") {

                datiFormattati = "preferenza(" + nome + ", " + componente + ").";

            } else {
                datiFormattati = "";
            }


        } else { //Caso 3: Disponibilità

            datiFormattati = "disponibilita(" + nome + ", " + componente + ").";
        }
        


    } else
        console.error("Tipo di dato non supportato:", typeof input);

    return datiFormattati;

}

    //Funzione per portare tutti i dati in una stringa per Clingo - STRING
function stringify(input) {
    let stringa = "";
    for (let i = 0; i < input.length; i++) {
        stringa += input[i].name + " " + input[i].turn + " " + input[i].preference + " " + input[i].disponibility + " ";
    }
    return stringa; 
}
