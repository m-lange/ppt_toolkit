/**
 * Lädt eine Einstellung hybrid (zuerst aus dem Dokument, dann aus dem LocalStorage).
 */
export const loadSetting = (key: string, defaultValue: any, isNumber: boolean = false): any => {
  let val: any = null;

  // 1. Priorität: Document Settings (aktuelle Präsentation)
  if (typeof Office !== 'undefined' && Office.context && Office.context.document) {
    val = Office.context.document.settings.get(key);
  }

  // 2. Priorität: LocalStorage (Gerätespezifisch), falls im Dokument nichts steht
  if (val === null || val === undefined || val === '') {
    const localVal = localStorage.getItem(`ppt_${key}`);
    if (localVal !== null && localVal !== '') {
      val = isNumber ? parseFloat(localVal) : localVal;
    }
  }

  // 3. Priorität: Fallback auf Standardwert
  if (val === null || val === undefined || val === '' || (isNumber && isNaN(val))) {
    val = defaultValue;
  }

  return val;
};

/**
 * Speichert eine Einstellung hybrid (sowohl im Dokument als auch im LocalStorage).
 */
export const saveSetting = (key: string, value: any): void => {
  // Im Dokument speichern
  if (typeof Office !== 'undefined' && Office.context && Office.context.document) {
    Office.context.document.settings.set(key, value);
    Office.context.document.settings.saveAsync();
  }
  // Auf dem Gerät speichern
  localStorage.setItem(`ppt_${key}`, String(value));
};