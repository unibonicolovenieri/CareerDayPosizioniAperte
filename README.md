# Career Day UniBO 2026 — Generatore DOCX

Script Node.js per generare un file Word strutturato con la lista delle aziende partecipanti al Career Day dell'Università di Bologna 2026, pensato per tracciare tirocini e posizioni aperte in ambito **Ingegneria Informatica**.

---

## Requisiti

- [Node.js](https://nodejs.org/) v14 o superiore
- Pacchetto npm `docx`

---

## Installazione

1. Installa Node.js dal sito ufficiale: https://nodejs.org/
2. Installa il pacchetto `docx` globalmente:

```bash
npm install -g docx
```

---

## Utilizzo

1. Posiziona lo script `create_docx.js` in una cartella a tua scelta
2. Apri il terminale in quella cartella ed esegui:

```bash
node create_docx.js
```

3. Troverai il file `career_day_unibo_2026_tirocini.docx` nella stessa cartella

---

## Struttura del file generato

Per ogni azienda viene creata una scheda con:

| Campo                      | Descrizione                                                                      |
| -------------------------- | -------------------------------------------------------------------------------- |
| **Header**           | Numero progressivo e nome azienda                                                |
| **URL Scheda**       | Link diretto alla pagina azienda sul sito Career Day                             |
| **Stato**            | Tre opzioni:`[ ] Da verificare` / `[X] Verificato` / `[NO] Non pertinente` |
| **Posizioni aperte** | Campo vuoto da compilare manualmente                                             |
| **Note**             | Campo libero per appunti                                                         |

---

## Come aggiornare le aziende

Le aziende sono definite nell'array `aziende` all'inizio dello script. Ogni oggetto ha tre campi:

```javascript
{ n: "001", nome: "Nome Azienda", id: "12345" }
```

| Campo    | Descrizione                                        |
| -------- | -------------------------------------------------- |
| `n`    | Numero progressivo a 3 cifre (es.`"001"`)        |
| `nome` | Nome dell'azienda come appare sul sito             |
| `id`   | ID numerico presente nell'URL della scheda azienda |

### Trovare l'ID di un'azienda

L'ID si trova nell'URL della pagina scheda sul sito Career Day:

```
https://eventi.unibo.it/careerday/aziende-partecipanti-2026/view/504292
                                                                  ^^^^^^
                                                                  questo è l'ID
```

### Aggiungere un'azienda

Aggiungi un nuovo oggetto all'array, rispettando l'ordine numerico:

```javascript
{ n: "176", nome: "Nuova Azienda S.p.A.", id: "99999" },
```

### Rimuovere un'azienda

Elimina la riga corrispondente dall'array e aggiorna i numeri progressivi se necessario.

---

## Struttura del progetto

```
📁 cartella-progetto/
├── create_docx.js                      ← script principale
├── README.md                           ← questo file
└── career_day_unibo_2026_tirocini.docx ← file generato (dopo l'esecuzione)
```

---

## Sito di riferimento

Le schede aziendali sono consultabili su:
https://eventi.unibo.it/careerday/aziende-partecipanti-2026
