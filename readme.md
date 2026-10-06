# Word in Markdown - Mermaid

Applicazione locale Python con interfaccia Streamlit per convertire documenti Word `.docx` o Markdown `.md` in diagrammi gerarchici Mermaid.

Permette di elaborare il documento completo oppure soltanto la sezione **Diagramma di Flusso**, visualizzare il risultato e scaricare Markdown, codice Mermaid e una pagina HTML con il diagramma.

## Requisiti

- Python con Streamlit compatibile con `st.mermaid_chart` e con il parametro `max_upload_size` del caricamento file. La versione verificata nel progetto e Streamlit **1.61.1**.
- `markdown-it-py`, usato per individuare i confini della sezione richiesta.
- Facoltativamente, `markitdown[docx]`, per la conversione Word principale.
- Facoltativamente, `magika`, per verificare il tipo reale dei documenti Word.
- Facoltativamente, Ollama con almeno un modello installato, raggiungibile su `http://localhost:11434`.

Per predisporre un ambiente Python dedicato su Windows:

```powershell
Set-Location C:\Progetto_AI\APP_Meremaid_Diagram
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install "streamlit>=1.61.1" markdown-it-py "markitdown[docx]" magika
```

Se le dipendenze sono gia presenti nell'interprete utilizzato, non occorre creare un nuovo ambiente. MarkItDown, Magika e Ollama non sono obbligatori: senza MarkItDown viene utilizzato il parser Word OOXML locale; senza Ollama resta disponibile il diagramma locale.

## Avvio

Con un ambiente Python che contiene le dipendenze:

```powershell
Set-Location C:\Progetto_AI\APP_Meremaid_Diagram
python -m streamlit run .\word_to_mermaid_llm_ultimate.py --server.address 127.0.0.1 --server.port 8502
```

Con l'ambiente virtuale creato sopra, sostituire `python` con `.\.venv\Scripts\python.exe`.

Aprire `http://127.0.0.1:8502/`. Per arrestare il server, premere `Ctrl+C` nel terminale.

La scheda **Word in Mermaid** del Launcher Dashboard avvia normalmente l'app sulla porta `8502`. Se la porta e occupata, usare una porta libera, ad esempio `--server.port 8503`, e aprire l'indirizzo corrispondente. La porta alternativa non modifica automaticamente quella configurata nel launcher.

Non avviare lo script con il solo comando `python word_to_mermaid_llm_ultimate.py`: l'interfaccia richiede il runtime Streamlit.

## Utilizzo

1. Caricare un documento `.docx` o `.md` tramite **Documento Word o Markdown**. Il limite e 20 MB.
2. Se Ollama e disponibile, scegliere il modello nella barra laterale.
3. Scegliere **Documento completo** oppure **Solo Diagramma di Flusso**.
4. Premere **Analizza documento** e attendere il completamento.
5. Consultare la scheda **Testo Markdown** e la scheda **Diagramma**. L'espansore **Codice Mermaid** mostra il sorgente del grafico.
6. Scaricare uno o piu dei tre formati disponibili.

Il cambio di file o di modalita elimina il risultato precedente: premere nuovamente **Analizza documento** per generare un risultato coerente con la nuova scelta.

## Modalita disponibili

### Documento completo

E la scelta predefinita e mantiene il comportamento originale. Il parser genera il diagramma dai titoli e dagli elenchi dell'intero documento convertito.

Per i Word, se e selezionato un modello Ollama, possono essere aggiunte al Markdown descrizioni delle immagini incorporate. La disponibilita e la qualita di queste descrizioni dipendono dal modello, che deve supportare gli input visivi.

### Solo Diagramma di Flusso

Elabora soltanto la sezione con titolo **Diagramma di Flusso**, senza distinguere maiuscole e minuscole. Sono riconosciuti anche titoli numerati, per esempio `3. Diagramma di Flusso`, e titoli autonomi in grassetto o testo semplice.

Per confini affidabili, usare gli stili Titolo/Heading in Word e titoli Markdown con `#`, `##` o `###` nei file `.md`.

- Con un titolo Markdown strutturato, la sezione termina al successivo titolo dello stesso livello o di livello superiore. Le sottosezioni di livello inferiore sono incluse.
- Con un titolo autonomo non strutturato, la sezione termina al successivo titolo strutturato o paragrafo interamente in grassetto.
- Se non viene trovato un confine successivo, la sezione prosegue fino alla fine del documento.
- Un titolo assente, duplicato o senza contenuto produce un errore: l'app non ripiega sulla conversione dell'intero documento.

In questa modalita non vengono aggiunte descrizioni globali delle immagini del Word. Anteprima e download Markdown contengono la sola sezione selezionata; Mermaid e HTML rappresentano la stessa selezione.

### Esempio Markdown

```markdown
# Procedura

## Introduzione
- Informazioni generali

## Diagramma di Flusso
- Ricezione richiesta
  - Controllo dei dati
  - Verifica documentazione
- Chiusura richiesta

### Dettagli
- Registrazione esito

## Note finali
- Archiviazione del documento
```

Con **Solo Diagramma di Flusso** vengono inclusi il titolo, i relativi elenchi e la sottosezione **Dettagli**. **Introduzione** e **Note finali** sono escluse.

## Come viene generato il diagramma

1. Un file `.md` viene decodificato in UTF-8; sono accettati anche file con BOM UTF-8.
2. Un Word viene controllato come archivio DOCX e, quando disponibile, identificato con Magika.
3. MarkItDown prova a convertire il Word in Markdown. In sua assenza, oppure in caso di errore del motore, viene usato il fallback OOXML locale.
4. Nella modalita mirata viene estratta la sezione richiesta tramite il parser Markdown.
5. Il parser Mermaid genera un grafo dall'alto verso il basso, con un nodo radice **Documento**. I titoli con `#` creano nodi gerarchici; gli elenchi puntati, numerati e annidati creano nodi collegati al rispettivo genitore. I blocchi di codice con flussi ASCII verticali vengono convertiti in sequenze e rami espliciti.
6. Se e selezionato un modello Ollama, viene richiesta una correzione della sintassi Mermaid. Sono accettate soltanto risposte che conservano nodi, etichette e collegamenti del diagramma locale. Una risposta non conforme o un errore di rete mantiene il diagramma locale e mostra un avviso.

Il risultato e una rappresentazione gerarchica: l'app non interpreta automaticamente condizioni, rami decisionali o sequenze operative descritte in prosa. I normali paragrafi e le tabelle possono comparire nel Markdown, ma non vengono trasformati in nodi dal parser del diagramma.

### Diagrammi testuali ASCII

Per rappresentare una sequenza operativa, scrivere il flusso ASCII su righe separate. Sono riconosciuti anche i normali paragrafi consecutivi di Word convertiti in Markdown, senza richiedere backtick. Un blocco di codice Markdown delimitato da tre backtick, oppure indentato di quattro spazi, resta il formato consigliato. Esempio:

````markdown
# 6. Diagramma di Flusso

```text
START
 |
 V
Controllo autorizzazioni
 |
 +--> KO --> Fine
 |
 V
Verifica periodi chiusi
 |
 V
FINE
```
````

La coppia di righe `|` e `V` collega il passaggio precedente al successivo. Una riga `+--> KO --> Fine` crea un ramo dal passaggio principale corrente senza interrompere la prosecuzione verticale. Le destinazioni finali `Fine`, `End` e `Stop` vengono condivise quando si ripetono, ignorando le maiuscole; le altre attivita ripetute restano nodi distinti.

Questo supporto riguarda sequenze verticali e rami laterali espliciti nel formato mostrato, non qualsiasi disegno ASCII. Sono accettate righe vuote tra i passaggi, come quelle prodotte dai paragrafi Word. Testo discorsivo e frecce incomplete non vengono interpretati come flussi ASCII. Separare il flusso da eventuale prosa circostante con titoli o un blocco di codice; verificare il testo estratto nella scheda **Testo Markdown**.

## Download

| Comando | File | Contenuto |
| --- | --- | --- |
| Scarica Markdown | `<nome>.md` | Markdown completo o sola sezione, secondo la modalita scelta |
| Scarica Mermaid | `<nome>.mmd` | Sorgente testuale del diagramma |
| Scarica HTML interattivo | `<nome>_diagramma.html` | Pagina HTML che visualizza il diagramma con Mermaid |

I nomi vengono sanificati a partire dal nome del documento caricato. L'app non sovrascrive il file originale.

L'HTML esportato carica **Mermaid 11 da jsDelivr**: richiede una connessione Internet al momento dell'apertura, salvo disponibilita della risorsa nella cache del browser. Non e una pagina completamente autonoma per l'uso offline.

## Limiti e gestione dei dati

- Sono supportati `.docx` e `.md`, non il vecchio formato `.doc`.
- I Word con oltre 5.000 elementi nell'archivio o piu di 100 MB di contenuto decompresso vengono rifiutati.
- I documenti vengono gestiti in memoria dall'app; i download vengono salvati dall'utente attraverso il browser.
- Il fallback Word conserva contenuti essenziali, titoli riconoscibili, elenchi, collegamenti e tabelle, ma non riproduce integralmente l'impaginazione originale.
- Per le descrizioni visive vengono considerate al massimo 20 immagini, fino a 8 MB ciascuna, nei formati PNG, JPEG, GIF e WebP.
- La richiesta di correzione a Ollama contiene al massimo i primi 50.000 caratteri del diagramma e ha un timeout di 90 secondi.
- Con il comando di avvio indicato, il server ascolta soltanto su `127.0.0.1`. Le richieste AI vengono inviate all'istanza locale di Ollama configurata nel codice.

## Problemi comuni

| Messaggio o problema | Come intervenire |
| --- | --- |
| Ollama non raggiungibile | Usare il parser locale oppure avviare Ollama e verificare `http://localhost:11434/api/tags` |
| Validazione Ollama non disponibile / timeout | Il diagramma locale resta utilizzabile; scegliere un modello piu leggero se necessario |
| Risposta LLM non conforme | L'app mantiene il diagramma locale senza usare la risposta del modello |
| Sezione non presente o non univoca | Inserire un solo titolo dedicato **Diagramma di Flusso**, preferibilmente con uno stile Heading in Word o `##` in Markdown |
| Sezione senza testo | Aggiungere contenuto dopo il titolo, prima della sezione successiva |
| Diagramma con pochi nodi | Usare titoli con `#`, elenchi oppure un flusso ASCII in un blocco di codice; il testo discorsivo non crea nodi |
| Markdown non leggibile | Salvare il file in UTF-8 |
| Errore relativo a `mermaid_chart` o `max_upload_size` | Aggiornare Streamlit nell'interprete usato per avviare l'app |
| Porta 8502 occupata | Arrestare l'istanza non necessaria oppure scegliere un'altra porta con `--server.port` |
| HTML senza diagramma | Verificare che il browser possa raggiungere il CDN jsDelivr |