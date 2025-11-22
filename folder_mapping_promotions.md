# Instrucciones: marcar “Promotions” en la columna Folder

1) En tu Apps Script abre el archivo de código y reemplaza la función `folderFromLabels_` por esta versión:
```javascript
function folderFromLabels_(labelIds) {
  if (!labelIds || !labelIds.length) return 'Other';

  // Categorías de Gmail
  if (labelIds.includes('CATEGORY_PROMOTIONS')) return 'Promotions';
  if (labelIds.includes('CATEGORY_SOCIAL')) return 'Social';
  if (labelIds.includes('CATEGORY_UPDATES')) return 'Updates';
  if (labelIds.includes('CATEGORY_FORUMS')) return 'Forums';

  // System labels
  if (labelIds.includes('SPAM'))  return 'Spam';
  if (labelIds.includes('TRASH')) return 'Trash';
  if (labelIds.includes('INBOX')) return 'Inbox';
  if (labelIds.includes('SENT'))  return 'Sent';
  if (labelIds.includes('DRAFT')) return 'Draft';

  return 'Other';
}
```

2) Guarda y despliega la versión web (Implementar → Nueva implementación o Actualizar la existente).

3) Ejecuta un `fullrescan` para recalcular la columna Folder en todos los mensajes (puedes lanzar con `run_fullrescan.ps1`).

Notas:
- Gmail añade la etiqueta `CATEGORY_PROMOTIONS` cuando el correo está en la pestaña Promociones.
- Si quieres ver todas las columnas recalculadas, limpia índice/estado (`clearHistoryAndIndex`) antes del full rescan.***
