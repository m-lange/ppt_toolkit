import React, { useState } from 'react';
import {
  FluentProvider,
  createLightTheme,
  Button,
  makeStyles,
  mergeClasses,
  tokens,
  type BrandVariants,
  type Theme
} from '@fluentui/react-components';
import { SettingsRegular } from '@fluentui/react-icons';
import { LayoutTab } from './components/LayoutTab';
import { SettingsTab } from './components/SettingsTab';
import { IconsTab } from './components/IconsTab';

const brandTheme: BrandVariants = {
  10: "#020401", 20: "#101C0F", 30: "#152F17", 40: "#173D1C",
  50: "#194B20", 60: "#1A5A25", 70: "#1A692A", 80: "#19782E",
  90: "#168833", 100: "#119938", 110: "#07A93C", 120: "#34B84E",
  130: "#60C46B", 140: "#82D087", 150: "#A2DCA3", 160: "#BFE7BF"
};

const powerPointTheme: Theme = {
  ...createLightTheme(brandTheme),
  fontSizeBase300: "12px",
};


// 2. Custom Styles für unsere eigene Button-Navigation
const useStyles = makeStyles({
  navBar: {
    display: 'flex',
    width: '100%',
    columnGap: '4px',
    alignItems: 'center',
  },
  navButton: {
    fontSize: '11px',
    fontWeight: 'normal !important', // Immer normal (nicht fett)
    paddingTop: '2px',
    paddingBottom: '2px',
    paddingLeft: '8px !important',
    paddingRight: '8px !important',
    minWidth: 'auto !important',
    minHeight: '26px',
    borderRadius: tokens.borderRadiusMedium,

    // 1. NORMAL: Dunkelgrauer Text, kein Hintergrund, unsichtbarer Rahmen
    color: `${tokens.colorNeutralForeground1} !important`, // Dunkelgrau/Fast Schwarz
    backgroundColor: 'transparent !important',
    border: '1px solid transparent !important', // Verhindert das "Zucken" beim Hovern

    // 2. HOVER (Unselected): Hellgrauer Hintergrund, mittelgrauer Rahmen
    '&:hover': {
      backgroundColor: `${tokens.colorNeutralBackground1Hover} !important`, // Hellgrau
      border: `1px solid ${tokens.colorNeutralStroke1} !important`, // Mittelgrauer Rahmen
      color: `${tokens.colorNeutralForeground1} !important`,
    },
  },

  // Diese Klasse wird nur angehängt, wenn der Button aktiv ist
  navButtonActive: {

    fontWeight: 'bold !important', // Immer normal (nicht fett)
    // 3. SELEKTIERT: Mittelgrauer Hintergrund, Rahmen & Text in Akzentfarbe
    backgroundColor: `${tokens.colorNeutralBackground4} !important`, // Mittelgrau
    border: `1px solid ${tokens.colorBrandStroke1} !important`, // Rahmen in Akzentfarbe (Orange)
    color: `${tokens.colorBrandForeground1} !important`, // Text in Akzentfarbe

    // Hover für SELECTED (damit es beim Drüberfahren nicht zurück auf das graue Hover springt)
    '&:hover': {
      backgroundColor: `${tokens.colorNeutralBackground3} !important`,
      border: `1px solid ${tokens.colorBrandStroke1} !important`,
      color: `${tokens.colorBrandForeground1} !important`,
    }
  },

  settingsButton: {
    marginLeft: 'auto',
  }
});


export const App: React.FC = () => {
  const classes = useStyles();
  const [selectedValue, setSelectedValue] = useState<string>('templates');

  const [reloadTrigger, setReloadTrigger] = useState<number>(0);

  const handleReloadIcons = () => {
    setReloadTrigger(prev => prev + 1);
  };


  return (
    <FluentProvider theme={powerPointTheme}>
      <div style={{ padding: '12px', display: 'flex', flexDirection: 'column', height: '100vh', boxSizing: 'border-box' }}>

        {/* Unsere eigene Navigationsleiste */}
        <div className={classes.navBar} role="tablist">

          <Button
            appearance="transparent"
            role="tab"
            // mergeClasses kombiniert die Standard-Klasse mit der Active-Klasse (falls ausgewählt)
            className={mergeClasses(
              classes.navButton,
              selectedValue === 'templates' && classes.navButtonActive
            )}
            onClick={() => setSelectedValue('templates')}
          >
            Layout
          </Button>

          <Button
            appearance="transparent"
            role="tab"
            className={mergeClasses(
              classes.navButton,
              selectedValue === 'icons' && classes.navButtonActive
            )}
            onClick={() => setSelectedValue('icons')}
          >
            Symbole
          </Button>

          <Button
            appearance="transparent"
            role="tab"
            icon={<SettingsRegular />}
            aria-label="Settings"
            className={mergeClasses(
              classes.navButton,
              classes.settingsButton,
              selectedValue === 'settings' && classes.navButtonActive
            )}
            onClick={() => setSelectedValue('settings')}
          />

        </div>

        {/* Content-Bereich */}
        <div style={{ flex: 1, marginTop: '16px', overflow: 'hidden' }}>
          {selectedValue === 'templates' && <LayoutTab />}
          {selectedValue === 'icons' && <IconsTab reloadTrigger={reloadTrigger} />}
          {selectedValue === 'settings' && <SettingsTab onReloadIcons={handleReloadIcons} />}
        </div>

      </div>
    </FluentProvider>
  );
};