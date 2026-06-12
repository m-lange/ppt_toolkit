import React, { useState, useEffect } from 'react';
import {
  SearchBox, Button, Field, makeStyles, mergeClasses, tokens, Text, Divider, Checkbox, Label
} from '@fluentui/react-components';
import {
  ShapesRegular, MaximizeRegular, CodeRegular, AddRegular,
  AlignLeftRegular, AlignCenterHorizontalRegular, AlignRightRegular,
  AlignTopRegular, AlignCenterVerticalRegular, AlignBottomRegular,
  AlignSpaceEvenlyHorizontalRegular, AlignSpaceEvenlyVerticalRegular // NEU: Icons fürs Verteilen
} from '@fluentui/react-icons';
import { insertBlueRectangle, showSlideDimensions, exportSelectedShapeToJson } from '../utils/powerpointApi';
import { SpinButton } from './SpinButton';
import { ToggleSwitch } from './ToggleSwitch';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '16px',
    height: '100%',
    overflowY: 'auto',
    paddingRight: '4px',
    paddingBottom: '16px'
  },
  section: { display: 'flex', flexDirection: 'column', gap: '6px' },
  heading: { fontSize: '11px', fontWeight: 'bold', color: tokens.colorNeutralForeground1 },
  divider: { marginTop: '2px', marginBottom: '4px' },

  alignmentGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(7, 32px)',
    gap: '6px',
  },
  gridBtn: {
    width: '32px',
    height: '32px',
    minWidth: '32px',
    padding: '0px',
    '& svg': { width: '20px', height: '20px' }
  },

  spaltenContent: {
    display: 'flex',
    gap: '16px',
    alignItems: 'flex-start',
    marginTop: '4px',
  },
  spaltenIcon: {
    width: '48px',
    height: '48px',
    flexShrink: 0,
    color: tokens.colorNeutralForeground2,
  },
  settingsColumn: {
    display: 'flex',
    flexDirection: 'column',
    gap: '6px',
    flex: 1,
  },
  settingRow: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
  },
  inputWrapper: {
    width: '85px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'flex-start',
  },
  checkboxAlign: {
    paddingLeft: '10px',
  },
  insertBtn: {
    marginTop: '4px',
    alignSelf: 'flex-end',
  }
});

// NEU: Jeder Button entscheidet selbst anhand der Anzahl (count), ob er aktiv ist!
const alignmentButtons = [
  // Zeile 1: Normale Ausrichtung (ab 1 Objekt aktiv)
  { id: 'left', icon: <AlignLeftRegular />, action: () => console.log('Linksbündig'), isEnabled: (count: number) => count >= 1 },
  { id: 'center', icon: <AlignCenterHorizontalRegular />, action: () => console.log('Zentriert'), isEnabled: (count: number) => count >= 1 },
  { id: 'right', icon: <AlignRightRegular />, action: () => console.log('Rechtsbündig'), isEnabled: (count: number) => count >= 1 },
  { id: 'space1', isSpace: true },
  { id: 'top', icon: <AlignTopRegular />, action: () => console.log('Oben ausrichten'), isEnabled: (count: number) => count >= 1 },
  { id: 'middle', icon: <AlignCenterVerticalRegular />, action: () => console.log('Mittig ausrichten'), isEnabled: (count: number) => count >= 1 },
  { id: 'bottom', icon: <AlignBottomRegular />, action: () => console.log('Unten ausrichten'), isEnabled: (count: number) => count >= 1 },

  // Zeile 2: Verteilen (erst ab 3 Objekten aktiv!)
  { id: 'distH', icon: <AlignSpaceEvenlyHorizontalRegular />, action: () => console.log('Horizontal verteilen'), isEnabled: (count: number) => count >= 3 },
  { id: 'distV', icon: <AlignSpaceEvenlyVerticalRegular />, action: () => console.log('Vertikal verteilen'), isEnabled: (count: number) => count >= 3 },
];

const customSvgIcon = `
  <svg width="48" height="48" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="3" y="3" width="18" height="18" rx="2" stroke="currentColor" stroke-width="1.5"/>
    <path d="M3 9H21" stroke="currentColor" stroke-width="1.5"/>
    <path d="M9 21V9" stroke="currentColor" stroke-width="1.5"/>
  </svg>
`;

const gridSvgIcon = `
  <svg width="48" height="48" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="3" y="3" width="18" height="18" rx="2" stroke="currentColor" stroke-width="1.5"/>
    <path d="M9 3V21" stroke="currentColor" stroke-width="1.5"/>
    <path d="M15 3V21" stroke="currentColor" stroke-width="1.5"/>
    <path d="M3 9H21" stroke="currentColor" stroke-width="1.5"/>
    <path d="M3 15H21" stroke="currentColor" stroke-width="1.5"/>
  </svg>
`;

const processSvgIcon = `
  <svg width="48" height="48" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
    <path d="M3 7H9L12 11.5L9 16H3L6 11.5L3 7Z" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/>
    <path d="M10 7H16L19 11.5L16 16H10L13 11.5L10 7Z" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/>
    <path d="M3 20H21" stroke="currentColor" stroke-width="1.5" stroke-dasharray="2 2"/>
  </svg>
`;

export const LayoutTab: React.FC = () => {
  const classes = useStyles();

  // NEU: Speichert jetzt die exakte Anzahl der markierten Objekte (0, 1, 2, 3...)
  const [selectedShapeCount, setSelectedShapeCount] = useState<number>(0);

  const [hasHeaders, setHasHeaders] = useState(true);
  const [columnCount, setColumnCount] = useState(3);

  const [gridColumns, setGridColumns] = useState(3);
  const [gridRows, setGridRows] = useState(3);
  const [gridHasHeaders, setGridHasHeaders] = useState(true);
  const [gridFirstColumnWidth, setGridFirstColumnWidth] = useState(0);

  const [processSteps, setProcessSteps] = useState(4);
  const [processSwimlanes, setProcessSwimlanes] = useState(1);
  const [processHasLabels, setProcessHasLabels] = useState(true);

  useEffect(() => {
    const checkSelection = async () => {
      if (typeof PowerPoint === 'undefined') return;
      try {
        await PowerPoint.run(async (context) => {
          const shapes = context.presentation.getSelectedShapes();
          shapes.load("items");
          await context.sync();
          // Speichert die exakte Anzahl
          setSelectedShapeCount(shapes.items.length);
        });
      } catch (error) {
        setSelectedShapeCount(0);
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

  const handleInsertCustomElement = () => {
    console.log(`Füge Spalten ein mit: Überschriften=${hasHeaders}, Spalten=${columnCount}`);
  };

  const handleInsertGrid = () => {
    console.log(`Füge Grid ein: Spalten=${gridColumns}, Zeilen=${gridRows}, Überschriften=${gridHasHeaders}, Erste Spalte=${gridFirstColumnWidth}cm`);
  };

  const handleInsertProcess = () => {
    console.log(`Füge Prozess ein: Schritte=${processSteps}, Swimlanes=${processSwimlanes}, Beschriftungen=${processHasLabels}`);
  };

  return (
    <div className={classes.container}>

      {/* --- BEREICH: AUSRICHTEN --- */}
      <div className={classes.section}>
        <div>
          <Text className={classes.heading}>Ausrichten</Text>
          <Divider className={classes.divider} />
        </div>

        <div className={classes.alignmentGrid}>
          {alignmentButtons.map((btn) => {
            if (btn.isSpace) return <div key={btn.id} />;
            return (
              <Button
                key={btn.id}
                icon={btn.icon}
                appearance="subtle"
                className={classes.gridBtn}
                onClick={btn.action}
                title={btn.id}
                // NEU: Ruft die isEnabled-Funktion des jeweiligen Buttons auf!
                disabled={!btn.isEnabled!(selectedShapeCount)}
              />
            );
          })}
        </div>
      </div>

      {/* --- BEREICH: SPALTEN --- */}
      <div className={classes.section}>
        <div>
          <Text className={classes.heading}>Spalten</Text>
          <Divider className={classes.divider} />
        </div>

        <div className={classes.spaltenContent}>
          <div
            className={classes.spaltenIcon}
            dangerouslySetInnerHTML={{ __html: customSvgIcon }}
          />

          <div className={classes.settingsColumn}>
            <div className={classes.settingRow}>
              <Label size="small">Anzahl Spalten</Label>
              <div className={classes.inputWrapper}>
                <SpinButton value={columnCount} onChange={setColumnCount} step={1} min={1} decimals={0} />
              </div>
            </div>
            <div className={classes.settingRow}>
              <Label size="small">Überschriften</Label>
              <ToggleSwitch checked={hasHeaders} onChange={setHasHeaders} />
            </div>
            <Button appearance="secondary" icon={<AddRegular />} onClick={handleInsertCustomElement} size="small" className={classes.insertBtn}>
              Einfügen
            </Button>
          </div>
        </div>
      </div>

      {/* --- BEREICH: GRID --- */}
      <div className={classes.section}>
        <div>
          <Text className={classes.heading}>Grid</Text>
          <Divider className={classes.divider} />
        </div>

        <div className={classes.spaltenContent}>
          <div
            className={classes.spaltenIcon}
            dangerouslySetInnerHTML={{ __html: gridSvgIcon }}
          />

          <div className={classes.settingsColumn}>
            <div className={classes.settingRow}>
              <Label size="small">Spalten</Label>
              <div className={classes.inputWrapper}>
                <SpinButton value={gridColumns} onChange={setGridColumns} step={1} min={1} decimals={0} />
              </div>
            </div>
            <div className={classes.settingRow}>
              <Label size="small">Zeilen</Label>
              <div className={classes.inputWrapper}>
                <SpinButton value={gridRows} onChange={setGridRows} step={1} min={1} decimals={0} />
              </div>
            </div>
            <div className={classes.settingRow}>
              <Label size="small">Erste Spalte</Label>
              <div className={classes.inputWrapper}>
                <SpinButton value={gridFirstColumnWidth} onChange={setGridFirstColumnWidth} step={0.01} min={0} decimals={2} decimalSeparator="," suffix=" cm" />
              </div>
            </div>
            <div className={classes.settingRow}>
              <Label size="small">Überschriften</Label>
              <ToggleSwitch checked={gridHasHeaders} onChange={setGridHasHeaders} />
            </div>
            <Button appearance="secondary" icon={<AddRegular />} onClick={handleInsertGrid} size="small" className={classes.insertBtn}>
              Einfügen
            </Button>
          </div>
        </div>
      </div>

      {/* --- BEREICH: PROZESS --- */}
      <div className={classes.section}>
        <div>
          <Text className={classes.heading}>Prozess</Text>
          <Divider className={classes.divider} />
        </div>

        <div className={classes.spaltenContent}>
          <div
            className={classes.spaltenIcon}
            dangerouslySetInnerHTML={{ __html: processSvgIcon }}
          />

          <div className={classes.settingsColumn}>
            <div className={classes.settingRow}>
              <Label size="small">Prozessschritte</Label>
              <div className={classes.inputWrapper}>
                <SpinButton value={processSteps} onChange={setProcessSteps} step={1} min={1} decimals={0} />
              </div>
            </div>
            <div className={classes.settingRow}>
              <Label size="small">Swimlane</Label>
              <div className={classes.inputWrapper}>
                <SpinButton value={processSwimlanes} onChange={setProcessSwimlanes} step={1} min={0} decimals={0} />
              </div>
            </div>
            <div className={classes.settingRow}>
              <Label size="small">Beschriftungen</Label>
              <ToggleSwitch checked={processHasLabels} onChange={setProcessHasLabels} />
            </div>
            <Button appearance="secondary" icon={<AddRegular />} onClick={handleInsertProcess} size="small" className={classes.insertBtn}>
              Einfügen
            </Button>
          </div>
        </div>
      </div>
    </div>
  );
};