import React, { useState, useEffect } from 'react';
import {
  makeStyles, mergeClasses, tokens, Text, Divider, Label, Input, Button, type InputOnChangeData
} from '@fluentui/react-components';
import { ArrowClockwiseRegular, TargetRegular } from '@fluentui/react-icons';
import { getMarginsFromSelection } from '../utils/powerpointApi';
import { SpinButton } from './SpinButton';
import { loadSetting, saveSetting } from '../utils/settings';

const useStyles = makeStyles({
  container: { display: 'flex', flexDirection: 'column', gap: '16px' },
  section: { display: 'flex', flexDirection: 'column', gap: '6px' },

  // NEU: Zusätzlicher Abstand nach oben für spezifische Bereiche
  extraTopMargin: { marginTop: '16px' },

  heading: { fontSize: '11px', fontWeight: 'bold', color: tokens.colorNeutralForeground1 },
  divider: { marginTop: '2px', marginBottom: '4px' },
  row: { display: 'flex', justifyContent: 'space-between', alignItems: 'center' },
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
  const [spacing, setSpacing] = useState({ horizontal: 0, vertical: 0 });
  const [iconUrl, setIconUrl] = useState('');

  useEffect(() => {
    const loadedMargins = {
      top: loadSetting('marginTop', 0, true),
      bottom: loadSetting('marginBottom', 0, true),
      left: loadSetting('marginLeft', 0, true),
      right: loadSetting('marginRight', 0, true),
    };
    setMargins(loadedMargins);

    const loadedSpacing = {
      horizontal: loadSetting('spacingHorizontal', 0, true),
      vertical: loadSetting('spacingVertical', 0, true),
    };
    setSpacing(loadedSpacing);

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

  const handleMarginChange = (side: keyof typeof margins, newValue: number) => {
    setMargins(prev => ({ ...prev, [side]: newValue }));
    saveSetting(`margin${side.charAt(0).toUpperCase() + side.slice(1)}`, newValue);
  };

  const handleSpacingChange = (side: keyof typeof spacing, newValue: number) => {
    setSpacing(prev => ({ ...prev, [side]: newValue }));
    saveSetting(`spacing${side.charAt(0).toUpperCase() + side.slice(1)}`, newValue);
  };

  const handleGetMarginsFromSelection = async () => {
    const newMargins = await getMarginsFromSelection();
    if (newMargins) {
      setMargins(newMargins);
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

      {/* --- BEREICH 1: RÄNDER --- */}
      <div className={classes.section}>
        <div>
          <Text className={classes.heading}>Ränder</Text>
          <Divider className={classes.divider} />
        </div>

        <div className={classes.row}>
          <Label size="small">Oben</Label>
          <SpinButton
            value={margins.top}
            onChange={(val) => handleMarginChange('top', val)}
            step={0.01}
            min={0}
            decimals={2}
            decimalSeparator=","
            suffix=" cm"
          />
        </div>

        <div className={classes.row}>
          <Label size="small">Unten</Label>
          <SpinButton
            value={margins.bottom}
            onChange={(val) => handleMarginChange('bottom', val)}
            step={0.01}
            min={0}
            decimals={2}
            decimalSeparator=","
            suffix=" cm"
          />
        </div>

        <div className={classes.row}>
          <Label size="small">Links</Label>
          <SpinButton
            value={margins.left}
            onChange={(val) => handleMarginChange('left', val)}
            step={0.01}
            min={0}
            decimals={2}
            decimalSeparator=","
            suffix=" cm"
          />
        </div>

        <div className={classes.row}>
          <Label size="small">Rechts</Label>
          <SpinButton
            value={margins.right}
            onChange={(val) => handleMarginChange('right', val)}
            step={0.01}
            min={0}
            decimals={2}
            decimalSeparator=","
            suffix=" cm"
          />
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

      {/* --- BEREICH 2: ABSTÄNDE --- */}
      <div className={classes.section}>
        <div>
          <Text className={classes.heading}>Abstände</Text>
          <Divider className={classes.divider} />
        </div>

        <div className={classes.row}>
          <Label size="small">Horizontal</Label>
          <SpinButton
            value={spacing.horizontal}
            onChange={(val) => handleSpacingChange('horizontal', val)}
            step={1}
            min={0}
            decimals={0}
            suffix=" pt"
          />
        </div>

        <div className={classes.row}>
          <Label size="small">Vertikal</Label>
          <SpinButton
            value={spacing.vertical}
            onChange={(val) => handleSpacingChange('vertical', val)}
            step={1}
            min={0}
            decimals={0}
            suffix=" pt"
          />
        </div>
      </div>

      {/* --- BEREICH 3: SYMBOLE (NEU: Mit extraTopMargin) --- */}
      <div className={mergeClasses(classes.section, classes.extraTopMargin)}>
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