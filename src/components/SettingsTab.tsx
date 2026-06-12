import React, { useState, useEffect } from 'react';
import {
  makeStyles, tokens, Text, Divider, Label, SpinButton, Input, Button,
  type SpinButtonChangeEvent, type SpinButtonOnChangeData, type InputOnChangeData
} from '@fluentui/react-components';
import { ArrowClockwiseRegular } from '@fluentui/react-icons';

const useStyles = makeStyles({
  container: { display: 'flex', flexDirection: 'column', gap: '16px' },
  section: { display: 'flex', flexDirection: 'column', gap: '6px' },
  heading: { fontSize: '11px', fontWeight: 'bold', color: tokens.colorNeutralForeground1 },
  divider: { marginTop: '2px', marginBottom: '4px' },
  row: { display: 'flex', justifyContent: 'space-between', alignItems: 'center' },

  spinContainer: { position: 'relative', width: '85px' },
  spinButton: { width: '100%' },
  cmSuffix: {
    position: 'absolute',
    right: '30px',
    top: '50%',
    transform: 'translateY(-50%)',
    fontSize: '11px',
    color: tokens.colorNeutralForeground3,
    pointerEvents: 'none',
  },

  inputContainer: { display: 'flex', flexDirection: 'column', gap: '4px' },
  urlInput: { width: '100%' },
  rightAlign: { display: 'flex', justifyContent: 'flex-end', marginTop: '4px' }
});

interface SettingsTabProps {
  onReloadIcons: () => void;
}

export const SettingsTab: React.FC<SettingsTabProps> = ({ onReloadIcons }) => {
  const classes = useStyles();

  // NEU: Ein Lade-Status, damit die Felder erst angezeigt werden, wenn die Daten da sind
  const [isLoaded, setIsLoaded] = useState(false);
  const [margins, setMargins] = useState({ top: 0, bottom: 0, left: 0, right: 0 });
  const [iconUrl, setIconUrl] = useState('');

  useEffect(() => {
    if (typeof Office !== 'undefined' && Office.context && Office.context.document) {
      const settings = Office.context.document.settings;
      setMargins({
        top: settings.get('marginTop') || 0,
        bottom: settings.get('marginBottom') || 0,
        left: settings.get('marginLeft') || 0,
        right: settings.get('marginRight') || 0,
      });
      setIconUrl(settings.get('iconUrl') || '');
    }
    // Daten sind geladen, UI darf gerendert werden
    setIsLoaded(true);
  }, []);

  const saveSetting = (key: string, value: any) => {
    if (typeof Office !== 'undefined' && Office.context && Office.context.document) {
      Office.context.document.settings.set(key, value);
      Office.context.document.settings.saveAsync();
    }
  };

  const handleMarginChange = (side: keyof typeof margins, _event: SpinButtonChangeEvent, data: SpinButtonOnChangeData) => {
    let val = data.value;

    // Wenn das Feld vom Nutzer geleert wird, gibt Fluent UI null oder undefined zurück.
    // In diesem Fall setzen wir den Wert sicherheitshalber auf 0.
    if (val === undefined || val === null) {
      val = 0;
    }

    // Jetzt weiß TypeScript sicher, dass 'val' eine Zahl ist
    if (typeof val === 'number' && !isNaN(val)) {
      setMargins(prev => ({ ...prev, [side]: val }));
      saveSetting(`margin${side.charAt(0).toUpperCase() + side.slice(1)}`, val);
    }
  };

  const handleUrlChange = (_event: React.ChangeEvent<HTMLInputElement>, data: InputOnChangeData) => {
    setIconUrl(data.value);
    saveSetting('iconUrl', data.value);
  };

  // NEU: Zeigt nichts an, bis die Werte aus PowerPoint geladen wurden
  if (!isLoaded) {
    return null;
  }

  return (
    <div className={classes.container}>

      {/* --- BEREICH 1: RÄNDER --- */}
      <div className={classes.section}>
        <div>
          <Text className={classes.heading}>Ränder</Text>
          <Divider className={classes.divider} />
        </div>

        <div className={classes.row}>
          <Label size="small">Oben</Label>
          <div className={classes.spinContainer}>
            <SpinButton
              defaultValue={margins.top} // WICHTIG: defaultValue statt value
              onChange={(e, d) => handleMarginChange('top', e, d)}
              className={classes.spinButton}
              appearance="filled-darker"
              size="small"
              step={0.01}
              min={0}
            />
            <span className={classes.cmSuffix}>cm</span>
          </div>
        </div>

        <div className={classes.row}>
          <Label size="small">Unten</Label>
          <div className={classes.spinContainer}>
            <SpinButton
              defaultValue={margins.bottom} // WICHTIG: defaultValue statt value
              onChange={(e, d) => handleMarginChange('bottom', e, d)}
              className={classes.spinButton}
              appearance="filled-darker"
              size="small"
              step={0.01}
              min={0}
            />
            <span className={classes.cmSuffix}>cm</span>
          </div>
        </div>

        <div className={classes.row}>
          <Label size="small">Links</Label>
          <div className={classes.spinContainer}>
            <SpinButton
              defaultValue={margins.left} // WICHTIG: defaultValue statt value
              onChange={(e, d) => handleMarginChange('left', e, d)}
              className={classes.spinButton}
              appearance="filled-darker"
              size="small"
              step={0.01}
              min={0}
            />
            <span className={classes.cmSuffix}>cm</span>
          </div>
        </div>

        <div className={classes.row}>
          <Label size="small">Rechts</Label>
          <div className={classes.spinContainer}>
            <SpinButton
              defaultValue={margins.right} // WICHTIG: defaultValue statt value
              onChange={(e, d) => handleMarginChange('right', e, d)}
              className={classes.spinButton}
              appearance="filled-darker"
              size="small"
              step={0.01}
              min={0}
            />
            <span className={classes.cmSuffix}>cm</span>
          </div>
        </div>
      </div>

      {/* --- BEREICH 2: SYMBOLE --- */}
      <div className={classes.section}>
        <div>
          <Text className={classes.heading}>Symbole</Text>
          <Divider className={classes.divider} />
        </div>

        <div className={classes.inputContainer}>
          <Label size="small">URL</Label>
          <Input
            value={iconUrl} // Bei normalen Textfeldern bleibt 'value' problemlos
            onChange={handleUrlChange}
            className={classes.urlInput}
            placeholder="https://..."
            appearance="filled-darker"
            size="small"
          />
        </div>

        <div className={classes.rightAlign}>
          <Button
            appearance="secondary"
            icon={<ArrowClockwiseRegular />}
            onClick={onReloadIcons}
            size="small"
          >
            Neu laden
          </Button>
        </div>
      </div>

    </div>
  );
};