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
import { TemplatesTab } from './components/TemplatesTab';
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
    fontWeight: 'bold',
    paddingTop: '2px',
    paddingBottom: '2px',
    paddingLeft: '12px',
    paddingRight: '12px',
    minHeight: '26px',
    borderRadius: tokens.borderRadiusMedium,

    // Standard-Zustand (Unselected)
    color: tokens.colorNeutralForeground2, // Grau
    backgroundColor: 'transparent',

    // Hover für UNSELECTED
    '&:hover': {
      color: `${tokens.colorNeutralForeground1} !important`, // Schwarz
      backgroundColor: 'transparent !important', // Kein grauer Hintergrund beim Hover
    },
  },

  // Diese Klasse wird nur angehängt, wenn der Button aktiv ist
  navButtonActive: {
    backgroundColor: `${tokens.colorBrandBackground2} !important`, // Heller Akzent
    color: `${tokens.colorBrandForeground1} !important`, // Akzentfarbe (Orange)

    // Hover für SELECTED (bleibt exakt gleich)
    '&:hover': {
      backgroundColor: `${tokens.colorBrandBackground2} !important`,
      color: `${tokens.colorBrandForeground1} !important`,
    }
  },

  settingsButton: {
    marginLeft: 'auto', // Drückt den Settings-Button nach rechts
  }
});

export const App: React.FC = () => {
  const classes = useStyles();
  const [selectedValue, setSelectedValue] = useState<string>('templates');

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
            Templates
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
          {selectedValue === 'templates' && <TemplatesTab />}
          {selectedValue === 'icons' && <IconsTab />}
          {selectedValue === 'settings' && <SettingsTab />}
        </div>

      </div>
    </FluentProvider>
  );
};