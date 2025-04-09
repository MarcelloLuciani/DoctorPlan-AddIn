// Funzione per eseguire Clingo con i dati letti dal foglio
async function eseguiClingoWasm(mappaDati) {
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

0 { assegna(C, pomeriggio, G) : turno(C, pomeriggio) } 1 :- giorno(G).
0 { assegna(C, mattina, G) : turno(C, mattina) } 1 :- giorno(G).

:- disponibilita(C, Max), Max+1 { assegna(C, _, _) }.
:- giorno(G), not 1 { assegna(_, _, G) }.

:~ assegna(C, M, G), preferenza(C, MP), M != MP. [1000@3, C, M, G]
#show assegna/3.
`;

        // Dispongo tutti i dati in un'unica stringa
        let stringaDati = stringify(mappaDati);

        // Creo il programma completo di Clingo
        const programmaCompleto = scriptClingo + "\n" + stringaDati;

        console.log("Programma completo di Clingo:", programmaCompleto);

        // Eseguo lo script
        const risultato = await clingo.run(programmaCompleto);
        console.log("Risultato di Clingo:", risultato);

        return risultato;
    } catch (error) {
        errorHandler(error);
    }
}