import React, { useState, useEffect } from 'react';
import {
  makeStyles, tokens, Text, Divider, Label, SpinButton, Input, Button,
  type SpinButtonChangeEvent, type SpinButtonOnChangeData, type InputOnChangeData
} from '@fluentui/react-components';
import { ArrowClockwiseRegular, TargetRegular } from '@fluentui/react-icons';
import { getMarginsFromSelection } from '../utils/powerpointApi';

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

  const [isLoaded, setIsLoaded] = useState(false);
  const [isSingleShapeSelected, setIsSingleShapeSelected] = useState(false);

  const [margins, setMargins] = useState({ top: 0, bottom: 0, left: 0, right: 0 });
  const [displayMargins, setDisplayMargins] = useState({ top: "0", bottom: "0", left: "0", right: "0" });
  const [iconUrl, setIconUrl] = useState('');

  const loadSetting = (key: string, defaultValue: any, isNumber: boolean = false) => {
    let val: any = null;
    if (typeof Office !== 'undefined' && Office.context && Office.context.document) {
      val = Office.context.document.settings.get(key);
    }
    if (val === null || val === undefined || val === '') {
      const localVal = localStorage.getItem(`ppt_${key}`);
      if (localVal !== null && localVal !== '') {
        val = isNumber ? parseFloat(localVal) : localVal;
      }
    }
    if (val === null || val === undefined || val === '' || (isNumber && isNaN(val))) {
      val = defaultValue;
    }
    return val;
  };

  const saveSetting = (key: string, value: any) => {
    if (typeof Office !== 'undefined' && Office.context && Office.context.document) {
      Office.context.document.settings.set(key, value);
      Office.context.document.settings.saveAsync();
    }
    localStorage.setItem(`ppt_${key}`, String(value));
  };

  useEffect(() => {
    const loadedMargins = {
      top: loadSetting('marginTop', 0, true),
      bottom: loadSetting('marginBottom', 0, true),
      left: loadSetting('marginLeft', 0, true),
      right: loadSetting('marginRight', 0, true),
    };

    setMargins(loadedMargins);
    setDisplayMargins({
      top: String(loadedMargins.top),
      bottom: String(loadedMargins.bottom),
      left: String(loadedMargins.left),
      right: String(loadedMargins.right),
    });

    setIconUrl(loadSetting('iconUrl', '', false));
    setIsLoaded(true);
  }, []);

  useEffect(() => {
    const checkSelection = async () => {
      if (typeof PowerPoint === 'undefined') return;
      try {
        await PowerPoint.run(async (context) => {
          const shapes = context.presentation.getSelectedShapes();
          shapes.load("items");
          await context.sync();
          setIsSingleShapeSelected(shapes.items.length === 1);
        });
      } catch (error) {
        setIsSingleShapeSelected(false);
      }
    };

    const handleSelectionChange = () => checkSelection();

    if (typeof Office !== 'undefined' && Office.context && Office.context.document) {
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handleSelectionChange);
      checkSelection();
    }

    return () => {
      if (typeof Office !== 'undefined' && Office.context && Office.context.document) {
        Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler: handleSelectionChange });
      }
    };
  }, []);

  // FIX: Unterscheidet zwischen Tippen und Klicken!
  const handleMarginChange = (side: keyof typeof margins, event: SpinButtonChangeEvent, data: SpinButtonOnChangeData) => {
    // 1. Wenn der Nutzer tippt (Event ist 'change')
    if (event.type === 'change') {
      if (data.displayValue !== undefined) {
        setDisplayMargins(prev => ({ ...prev, [side]: data.displayValue! }));
      }
    }
    // 2. Wenn der Nutzer auf die Pfeile klickt oder Pfeiltasten nutzt
    else {
      if (data.value !== undefined && data.value !== null && !isNaN(data.value)) {
        setMargins(prev => ({ ...prev, [side]: data.value! }));
        setDisplayMargins(prev => ({ ...prev, [side]: String(data.value!) }));
        saveSetting(`margin${side.charAt(0).toUpperCase() + side.slice(1)}`, data.value);
      }
    }
  };

  // FIX: Speichert die getippte Zahl, wenn man das Feld verlässt
  const handleMarginBlur = (side: keyof typeof margins) => {
    let textValue = displayMargins[side].trim();
    if (textValue === '') textValue = '0'; // Leeres Feld wird zu 0

    // Ersetzt Komma durch Punkt (für deutsche Tastaturen) und wandelt in Zahl um
    const parsed = parseFloat(textValue.replace(',', '.'));

    if (!isNaN(parsed)) {
      setMargins(prev => ({ ...prev, [side]: parsed }));
      setDisplayMargins(prev => ({ ...prev, [side]: String(parsed) }));
      saveSetting(`margin${side.charAt(0).toUpperCase() + side.slice(1)}`, parsed);
    } else {
      // Falls Quatsch eingetippt wurde, auf letzten gültigen Wert zurücksetzen
      setDisplayMargins(prev => ({ ...prev, [side]: String(margins[side]) }));
    }
  };

  const handleGetMarginsFromSelection = async () => {
    const newMargins = await getMarginsFromSelection();
    if (newMargins) {
      setMargins(newMargins);
      setDisplayMargins({
        top: String(newMargins.top),
        bottom: String(newMargins.bottom),
        left: String(newMargins.left),
        right: String(newMargins.right)
      });

      saveSetting('marginTop', newMargins.top);
      saveSetting('marginBottom', newMargins.bottom);
      saveSetting('marginLeft', newMargins.left);
      saveSetting('marginRight', newMargins.right);
    }
  };

  const handleUrlChange = (_event: React.ChangeEvent<HTMLInputElement>, data: InputOnChangeData) => {
    setIconUrl(data.value);
    saveSetting('iconUrl', data.value);
  };

  if (!isLoaded) return null;

  return (
    <div className={classes.container}>

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
              displayValue={displayMargins.top}
              onChange={(e, d) => handleMarginChange('top', e, d)}
              onBlur={() => handleMarginBlur('top')}
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
              value={margins.bottom}
              displayValue={displayMargins.bottom}
              onChange={(e, d) => handleMarginChange('bottom', e, d)}
              onBlur={() => handleMarginBlur('bottom')}
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
              value={margins.left}
              displayValue={displayMargins.left}
              onChange={(e, d) => handleMarginChange('left', e, d)}
              onBlur={() => handleMarginBlur('left')}
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
              value={margins.right}
              displayValue={displayMargins.right}
              onChange={(e, d) => handleMarginChange('right', e, d)}
              onBlur={() => handleMarginBlur('right')}
              className={classes.spinButton}
              appearance="filled-darker"
              size="small"
              step={0.01}
              min={0}
            />
            <span className={classes.cmSuffix}>cm</span>
          </div>
        </div>

        <div className={classes.rightAlign}>
          <Button
            appearance="secondary"
            icon={<TargetRegular />}
            onClick={handleGetMarginsFromSelection}
            size="small"
            disabled={!isSingleShapeSelected}
          >
            Von Auswahl übernehmen
          </Button>
        </div>
      </div>

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