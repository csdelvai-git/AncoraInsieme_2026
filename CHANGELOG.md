# Ancora Insieme — DdT App · Changelog

**Progetto:** Torneo aziendale di calcetto BLM Group dedicato ai colleghi scomparsi  
**Tecnologia:** HTML/JS monolitico, pubblicazione su GitHub Pages, dati su `giocatori.json`  
**Autore:** Claudio S. Delvai — ancorainsieme.ddt@gmail.com  

---

## v8.7 — 26/07/2026
**Classifica pubblica con colonne ordinabili**
- Click sull'intestazione di colonna → ordina per quella colonna
- Click di nuovo → inverte l'ordine (▲ / ▼)
- Colonne ordinabili: Nome, ±, ELO, Gare, Vittorie, Win%, Diff Gol, Ultime 5, Idx Partner%, Idx Rivale%
- Default invariato: ELO decrescente
- I dati della classifica sono ora serializzati come JSON e renderizzati lato client

---

## v8.6 — 26/07/2026
**Grafico ELO interattivo con selezione multipla**
- Tutti i giocatori nel grafico (rimosso filtro Top N)
- Click su un nome in legenda → evidenzia la curva, attenua le altre
- Selezione multipla fino a 3 giocatori contemporaneamente
- Click su giocatore già selezionato → deseleziona
- Curve che partono dalla prima giornata reale (niente interpolazione all'indietro)
- Opacità attenuati calibrata per mantenere leggibilità

**Fix link navigazione**
- Home dalla classifica: corretto path `../../index.html`
- Classifica → grafico: corretto path `../../Grafici/2026/...`

---

## v8.5 — 24/07/2026
**Fix pubblicazione schede giocatori**
- Rimosso download locale delle schede individuali (troppo numerose, inutili offline)
- Solo pubblicazione su GitHub Pages

---

## v8.4 — 24/07/2026
**Ristrutturazione cartelle GitHub Pages + schede giocatori individuali**

Nuova struttura path:
- `Classifiche/2026/Classifica_{Sede}.html`
- `Storici/2026/Storico_{Sede}.html`
- `Storici/2026/Storico_Completo.html`
- `Grafici/2026/Grafico_Giornate_{Sede}.html`
- `Giocatori/2026/Scheda_{Nome}.html` ← nuovo
- `index.html` → rimane in root

Nuova funzione `_generaHTMLSchedaGiocatore(nome)`:
- Pagina individuale per ogni giocatore attivo (gare > 0)
- KPI: ELO, gare, vittorie, win%, diff gol
- Pallini ultime 5 partite
- Tabelle compagni e avversari con soglie evidenziate (⚠️ se sforata)
- Link Home e Classifica sede

Nome giocatore nella classifica pubblica → link cliccabile alla scheda individuale  
Tutte le pagine in sottocartella usano path relativi `../../index.html` per Home  

---

## v8.3 — 23/07/2026
**Semplificazione mail di iscrizione**
- Un solo link alla landing page `index.html` invece di elencare classifica/storico/regolamento/edicola separatamente
- Aggiunto link separato al form segnalazione risultati

---

## v8.2 — 23/07/2026
**Fix leggibilità numeri ELO**
- Font cambiato da `Barlow Condensed` a `Barlow` per i numeri ELO (il "6" risultava confondibile con "8")
- Aggiunto `font-variant-numeric: tabular-nums` e `letter-spacing` nell'app (classe `.elo-number`) e negli HTML esportati

---

## v8.1 — 23/07/2026
**Navigazione bidirezionale completa**
- Tutte le pagine pubblicate linkano `🏠 Home`
- Navigazione: index ↔ classifica ↔ grafico ↔ storico

---

## v8.0 — 23/07/2026
**Box sostenitori in classifica**
- Aggiunto box azzurro "❤️ Grazie ai nostri sostenitori" in fondo alla classifica pubblica
- Elenca i giocatori iscritti ma a 0 gare, con testo motivazionale
- Appare solo se ci sono sostenitori; scompare automaticamente quando giocano la prima partita

---

## v7.9 — 23/07/2026
**Nota sostenitori dinamica**
- La nota in fondo alla classifica conta dinamicamente i sostenitori presenti

---

## v7.8 — 23/07/2026
**Nota sostenitori in classifica pubblica**
- Aggiunta nota statica in fondo: "In classifica compaiono solo i giocatori che hanno disputato almeno una partita..."

---

## v7.7 — 23/07/2026
**Filtro sostenitori (giocatori a 0 gare)**
- Approccio implicito: i giocatori con 0 partite vengono esclusi da tutti gli output pubblici automaticamente
- Classifica HTML, Top N grafico, KPI sotto-soglia/max-raggiunto filtrati con `gare > 0`
- Reversibile automaticamente alla prima partita giocata
- Nessun flag esplicito necessario

---

## v7.6 — 23/07/2026
**Visibilità parametri ELO nei dialog GitHub**
- Aggiunta funzione `_riassuntoParametri()` 
- Dialog carica/salva GitHub mostra i parametri per sede (pesi forte/debole, soglie compagno/avversario)
- I due DdT possono confrontare i parametri prima di sovrascrivere

---

## v7.5 — 23/07/2026
**Landing page `index.html` generata automaticamente**
- Generata e pubblicata da "Pubblica HTML" insieme agli altri file
- Card per sede con link a Classifica, Storico Partite, Andamento ELO
- Sezione Archivio Completo (Storico_Completo.html)
- Sezione Edicola, Regolamento PDF, Form segnalazione, PayPal
- Stile coerente con le classifiche (sfondo bianco, header verde scuro con logo)

---

## v7.4 — 23/07/2026
**Pulizia funzioni ridondanti**
- Rimossa funzione `esportaHTMLClassifica()` e relativi bottoni (toolbar + tab Classifica)
- Ridondante con la pubblicazione automatica via "Pubblica HTML"

---

## v7.3 — 23/07/2026
**Grafico ELO per giornata generato automaticamente**
- Aggiunto parametro configurabile "Top N giocatori da esportare" (default 10, range 1-50)
- "Pubblica HTML" genera automaticamente il grafico ELO per giornata per ogni sede
- Nuova funzione `_generaHTMLGraficoDatePubblico()`: sfondo bianco, altezza 600px, palette colori distinti
- Link di navigazione verso la classifica
- Didascalia dinamica sopra il canvas con spiegazione dell'asse orizzontale

---

## v7.2 e precedenti — sessioni precedenti

Le versioni da v4.6 a v7.2 sono documentate nella sessione di sviluppo precedente e includono:

- **v4.6–v5.x**: gestione anagrafica giocatori (sposta sede, rinomina, flag enclave), pannello mail iscrizione, storico gare multi-sede con filtri
- **v5.x–v6.x**: import CSV con gestione doppioni, deduzione automatica sede partita, tab "Stat Giornate" con grafico ELO temporale
- **v6.x–v7.x**: pesi ELO asimmetrici forte/debole (per mitigare "ghettizzazione" giocatori deboli), pannello dettaglio giocatore con tabelle compagni/avversari e soglie configurabili, frecce posizione e pallini ultime 5 partite in classifica pubblica, download automatico locale ad ogni pubblicazione
- **v7.0–v7.2**: stabilizzazione, fix bug vari, verifica sintassi JS sistematica con `node --check`

---

## File correlati

| File | Versione | Descrizione |
|------|----------|-------------|
| `BLM_Torneo_v8_7.html` | v8.7 | App DdT — versione corrente |
| `AncoraInsieme_Form.html` | v2.8+ | Form segnalazione risultati (con pre-compilazione URL) |
| `AncoraInsieme_Sorteggio.html` | v1.2 | App sorteggio partite serata |
| `AncoraInsieme.gs` | v4.5 | Script Google Sheet — gestione segnalazioni |
| `AncoraInsieme_Regolamento.pdf` | — | Regolamento ufficiale torneo |

---

## Google Apps Script — Changelog

### v4.5 — 26/07/2026
**Criteri doppione rivisti e performance**

Nuovi criteri:
- 4/4 giocatori uguali → doppione (con o senza stesso punteggio)
- 3/4 giocatori uguali + stesso punteggio → probabile doppione (indica il nome diverso)

Ottimizzazioni performance:
- `_controllaDoppioni()` ora accetta cache dati in memoria (zero `getRange()` nel loop)
- `ricalcolaDoppioni()` legge tutto il foglio una sola volta e passa la cache
- `controllaSegnalazione()` unico `flush()` finale invece di uno per riga
- Partite Validate/Esportate usate come riferimento ma non riflaggiate
- Partite "Da confermare" saltate nel confronto tra loro

### v4.4 e precedenti
- Setup fogli, trigger form, notifiche email DdT e segnalante
- Validazione risultato (vittoria a 10 gol, scarto minimo 2, no pareggi)
- Controllo doppioni base (finestra 15 minuti configurabile)
- Esportazione CSV partite validate con spostamento in foglio Esportate
- Menu DdT con validazione, marcatura "Da confermare", rimessa in attesa
