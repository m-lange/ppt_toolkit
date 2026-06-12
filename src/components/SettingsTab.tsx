import React, { useState, useEffect } from 'react';
import {
  makeStyles, tokens, Text, Divider, Label, SpinButton, Input, Button,
  type SpinButtonChangeEvent, type SpinButtonOnChangeData, type InputOnChangeData
} from '@fluentui/react-components';
import { ArrowClockwiseRegular } from '@fluentui/react-icons';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '16px',
  },
  section: {
    display: 'flex',
    flexDirection: 'column',
    gap: '6px',
  },
  heading: {
    fontSize: '11px',
    fontWeight: 'bold',
    color: tokens.colorNeutralForeground1,
  },
  divider: {
    marginTop: '2px',
    marginBottom: '4px',
  },
  row: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
  },

  // --- CSS Trick für das "cm" im Input ---
  spinContainer: {
    position: 'relative',
    width: '85px', // NEU: Etwas breiter gemacht (vorher 75px)
  },
  spinButton: {
    width: '100%',
  },
  cmSuffix: {
    position: 'absolute',
    right: '30px', // NEU: Weiter nach links verschoben (vorher 24px)
    top: '50%',
    transform: 'translateY(-50%)',
    fontSize: '11px',
    color: tokens.colorNeutralForeground3,
    pointerEvents: 'none',
  },
  // ---------------------------------------

  inputContainer: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
  },
  urlInput: {
    width: '100%',
  },
  rightAlign: {
    display: 'flex',
    justifyContent: 'flex-end',
    marginTop: '4px',
  }
});

interface SettingsTabProps {
  onReloadIcons: () => void;
}

export const SettingsTab: React.FC<SettingsTabProps> = ({ onReloadIcons }) => {
  const classes = useStyles();
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
  }, []);

  const saveSetting = (key: string, value: any) => {
    if (typeof Office !== 'undefined' && Office.context && Office.context.document) {
      Office.context.document.settings.set(key, value);
      Office.context.document.settings.saveAsync();
    }
  };

  const handleMarginChange = (side: keyof typeof margins, _event: SpinButtonChangeEvent, data: SpinButtonOnChangeData) => {
    const val = data.value !== undefined ? data.value : 0;
    setMargins(prev => ({ ...prev, [side]: val }));
    saveSetting(`margin${side.charAt(0).toUpperCase() + side.slice(1)}`, val);
  };

  const handleUrlChange = (_event: React.ChangeEvent<HTMLInputElement>, data: InputOnChangeData) => {
    setIconUrl(data.value);
    saveSetting('iconUrl', data.value);
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
          <Label size="small">Oben</Label>
          <div className={classes.spinContainer}>
            <SpinButton
              value={margins.top}
              onChange={(e, d) => handleMarginChange('top', e, d)}
              className={classes.spinButton}
              appearance="filled-lighter" // NEU: Gefüllter, heller Hintergrund
              size="small"                // NEU: Kompaktere Größe
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
              value={margins.bottom}
              onChange={(e, d) => handleMarginChange('bottom', e, d)}
              className={classes.spinButton}
              appearance="filled-lighter"
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
              value={margins.left}
              onChange={(e, d) => handleMarginChange('left', e, d)}
              className={classes.spinButton}
              appearance="filled-lighter"
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
              value={margins.right}
              onChange={(e, d) => handleMarginChange('right', e, d)}
              className={classes.spinButton}
              appearance="filled-lighter"
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
            value={iconUrl}
            onChange={handleUrlChange}
            className={classes.urlInput}
            placeholder="https://..."
            appearance="filled-lighter" // NEU: Gefüllter, heller Hintergrund
            size="small"                // NEU: Kompaktere Größe
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