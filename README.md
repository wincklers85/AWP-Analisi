# AWP Analytics – versione Render

## Funzioni
- Import manuale di un singolo `Sinottico*.xlsx`.
- Import manuale di archivi `Sinottico*.zip` contenenti uno o più `Sinottico*.xlsx`.
- Import da repository GitHub pubblica o privata.
- Vengono considerati esclusivamente i file il cui nome inizia con `Sinottico`.
- Controllo duplicati sia sul file completo sia sulle singole letture macchina.
- Database SQLite persistente nella cartella `instance`.
- Pagina **Analizza** per macchine con ultima posizione `Attiva` e `E/M = E`.
- Semaforo statistico verde/arancione/rosso basato su scostamento payout, posizione ciclo e attività storica.

## Pubblicazione su Render
1. Carica questa cartella in una repository GitHub.
2. Su Render scegli **New > Blueprint** e collega la repository.
3. Render leggerà `render.yaml` e creerà servizio e disco persistente.
4. Il piano `starter` è indicato perché SQLite richiede un disco persistente. Senza disco i dati possono essere persi a ogni nuova distribuzione o riavvio.
5. Per importare da una repository GitHub privata, configura `GITHUB_TOKEN` nelle variabili ambiente Render con un token che abbia accesso in lettura alla repository.

## Import da GitHub
Nella pagina iniziale inserisci:
- URL repository, per esempio `https://github.com/utente/sinottici`;
- branch, normalmente `main`;
- cartella opzionale, per esempio `archivio/2026`.

Il programma esamina anche le sottocartelle e scarica soltanto `.xlsx` o `.zip` che iniziano con `Sinottico`.

## Avvio locale
```powershell
cd C:\awp_analytics
py -m venv venv
.\venv\Scripts\activate
py -m pip install -r requirements.txt
py app.py
```
Apri `http://127.0.0.1:5000`.

## Nota sull'indicatore Analizza
La classifica descrive una propensione statistica ricavata dai contatori e dallo storico. Non costituisce previsione certa né garanzia di vincita.

## Funzioni aggiunte - Salute del parco

- Pagina **Salute del parco** con punteggio 0-100 per ogni macchina attiva e in esercizio.
- Il punteggio combina freschezza della lettura, continuità dei contatori, rendimento rispetto allo stesso modello e completezza ciclo/payout.
- Classifica priorità con semaforo verde, arancione e rosso.
- Raccomandazioni operative per fermo contatori, possibile gettoniera/refill/hopper e rendimento sotto media.
- Scheda macchina estesa con media giocato, confronto con modello, timeline spostamenti e anomalie storiche.
- Gli indicatori sono gestionali e statistici: non garantiscono vincite future.

## Analizza – indice statistico v3
La pagina Analizza assegna a ogni macchina attiva e in esercizio un indice gestionale 0–100. Il calcolo usa scostamento dal payout teorico, posizione stimata nel ciclo, continuità di gioco, attività rispetto allo stesso modello, freschezza della lettura e quantità di storico disponibile. Ogni riga mostra motivazioni, criticità, affidabilità dei dati e dettaglio dei contributi al punteggio. L'indice non è una probabilità matematica e non garantisce un pagamento.
