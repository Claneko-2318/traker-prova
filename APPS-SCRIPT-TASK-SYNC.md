# Sincronizzazione task tra PC e cellulare

Le sessioni erano gia' condivise tramite Google Sheets; le task, invece, erano salvate soltanto nel browser del singolo dispositivo.

Nel progetto Apps Script collegato al tracker, integra questi due casi nello stesso `doPost` / `doGet` gia' usato per sessioni, categorie e progetti.

## Nel ramo `doPost`

```javascript
if (data.action === 'saveTasks') {
  saveTasks_(data.tasks || []);
  return json_({ok: true});
}
```

## Nella risposta dell'azione `get`

Aggiungi la proprieta' `tasks` all'oggetto JSON restituito:

```javascript
tasks: getTasks_()
```

## Funzioni da aggiungere a `Code.gs`

```javascript
function saveTasks_(tasks) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Tasks');
  if (!sh) sh = ss.insertSheet('Tasks');

  sh.clearContents();
  sh.getRange(1, 1, 1, 2).setValues([['id', 'json']]);
  if (!Array.isArray(tasks) || !tasks.length) return;

  const rows = tasks.map(function (task) {
    return [String(task.id || ''), JSON.stringify(task)];
  });
  sh.getRange(2, 1, rows.length, 2).setValues(rows);
}

function getTasks_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
  if (!sh || sh.getLastRow() < 2) return [];

  return sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues()
    .map(function (row) {
      try { return JSON.parse(row[0]); } catch (e) { return null; }
    })
    .filter(function (task) { return task && task.id; });
}
```

Dopo la modifica, crea una **nuova distribuzione** della Web App mantenendo lo stesso accesso pubblico. Poi apri/aggiorna prima il PC che contiene le task: se il foglio `Tasks` e' vuoto, il tracker migrera' automaticamente l'elenco locale senza cancellarlo. Dopo questa prima sincronizzazione, apri il tracker sul cellulare.
