import React, { useState, useEffect } from 'react';
import {
  makeStyles,
  tokens,
  Text,
  Divider,
  Label,
  SpinButton,
  Input,
  Button,
  type SpinButtonChangeEvent,
  type SpinButtonOnChangeData,
  type InputOnChangeData
} from '@fluentui/react-components';
import { ArrowClockwiseRegular } from '@fluentui/react-icons';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '24px', // Abstand zwischen den beiden großen Bereichen
    padding: '4px 0',
  },
  section: {
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
  },
  // Formatierung für die Überschriften (11px, fett)
  heading: {
    fontSize: '11px',
    fontWeight: 'bold',
    color: tokens.colorNeutralForeground1,
  },
  divider: {
    marginTop: '4px',
  },
  // Layout für die Ränder (Label links, Input rechts)
  row: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
  },
  spinContainer: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  spinButton: {
    width: '70px', // Kompakte Breite für den Spinner
  },
  // Layout für die Symbole (Label oben, Input 100% Breite)
  inputContainer: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
  },
  urlInput: {
    width: '100%',
  },
  // Button rechtsbündig
  rightAlign: {
    display: 'flex',
    justifyContent: 'flex-end',
  }
});

interface SettingsTabProps {
  onReloadIcons: () => void;
}

export const SettingsTab: React.FC<SettingsTabProps> = ({ onReloadIcons }) => {
  const classes = useStyles();

  // State für die Eingabefelder
  const [margins, setMargins] = useState({ top: 0, bottom: 0, left: 0, right: 0 });
  const [iconUrl, setIconUrl] = useState('');

  // 1. Beim Starten des Add-Ins die gespeicherten Settings laden
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
  }, []);

  // 2. Hilfsfunktion zum sofortigen Speichern in der PowerPoint-Datei
  const saveSetting = (key: string, value: any) => {
    if (typeof Office !== 'undefined' && Office.context && Office.context.document) {
      Office.context.document.settings.set(key, value);
      Office.context.document.settings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error("Fehler beim Speichern der Einstellung:", result.error.message);
        }
      });
    }
  };

  // 3. Handler für die Spinner-Buttons (Ränder)
  const handleMarginChange = (side: keyof typeof margins, _event: SpinButtonChangeEvent, data: SpinButtonOnChangeData) => {
    if (data.value !== undefined) {
      setMargins(prev => ({ ...prev, [side]: data.value! }));

      // Baut den Key zusammen (z.B. "marginTop", "marginLeft") und speichert ihn
      const key = `margin${side.charAt(0).toUpperCase() + side.slice(1)}`;
      saveSetting(key, data.value);
    }
  };

  // 4. Handler für das URL-Feld
  const handleUrlChange = (_event: React.ChangeEvent<HTMLInputElement>, data: InputOnChangeData) => {
    setIconUrl(data.value);
    saveSetting('iconUrl', data.value);
  };

  // 5. Handler für den "Neu laden" Button
  const handleReload = () => {
    console.log("Lade Icons neu von URL:", iconUrl);
    // Hier kannst du später die Logik einbauen, um die IconsTab.tsx anzuweisen, die JSON neu zu fetchen
    onReloadIcons();
  };

  return (
    <div className={classes.container}>

      {/* --- BEREICH 1: RÄNDER --- */}
      <div className={classes.section}>
        <div>
          <Text className={classes.heading}>Ränder</Text>
          <Divider className={classes.divider} />
        </div>

        <div className={classes.row}>
          <Label>Oben</Label>
          <div className={classes.spinContainer}>
            <SpinButton value={margins.top} onChange={(e, d) => handleMarginChange('top', e, d)} className={classes.spinButton} />
            <Text size={200}>cm</Text>
          </div>
        </div>

        <div className={classes.row}>
          <Label>Unten</Label>
          <div className={classes.spinContainer}>
            <SpinButton value={margins.bottom} onChange={(e, d) => handleMarginChange('bottom', e, d)} className={classes.spinButton} />
            <Text size={200}>cm</Text>
          </div>
        </div>

        <div className={classes.row}>
          <Label>Links</Label>
          <div className={classes.spinContainer}>
            <SpinButton value={margins.left} onChange={(e, d) => handleMarginChange('left', e, d)} className={classes.spinButton} />
            <Text size={200}>cm</Text>
          </div>
        </div>

        <div className={classes.row}>
          <Label>Rechts</Label>
          <div className={classes.spinContainer}>
            <SpinButton value={margins.right} onChange={(e, d) => handleMarginChange('right', e, d)} className={classes.spinButton} />
            <Text size={200}>cm</Text>
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
          <Label>URL</Label>
          <Input
            value={iconUrl}
            onChange={handleUrlChange}
            className={classes.urlInput}
            placeholder="https://..."
          />
        </div>

        <div className={classes.rightAlign}>
          <Button
            appearance="secondary"
            icon={<ArrowClockwiseRegular />}
            onClick={handleReload}
            size="small"
          >
            Neu laden
          </Button>
        </div>
      </div>

    </div>
  );
};