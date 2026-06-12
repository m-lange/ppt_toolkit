import React, { useState } from 'react';
import {
  FluentProvider,
  createLightTheme,
  TabList,
  Tab,
  makeStyles,
  tokens,
  type BrandVariants,
  type Theme,
  type SelectTabData,
  type SelectTabEvent
} from '@fluentui/react-components';
import { SettingsRegular, ImageRegular } from '@fluentui/react-icons';
import { TemplatesTab } from './components/TemplatesTab';
import { SettingsTab } from './components/SettingsTab';
import { IconsTab } from './components/IconsTab';

// 1. PowerPoint Brand Colors (Orange/Rot)
const powerPointBrand: BrandVariants = {
  10: "#040100",
  20: "#1F0802",
  30: "#350B02",
  40: "#480D02",
  50: "#5C1002",
  60: "#701202",
  70: "#851502",
  80: "#9B1702",
  90: "#B21A02",
  100: "#C43E1C", // PowerPoint Hauptfarbe
  110: "#D05232",
  120: "#DB6547",
  130: "#E4785D",
  140: "#EC8B73",
  150: "#F39E8A",
  160: "#F8B1A1",
};

// 2. Theme anpassen: Globale Schriftgröße auf 12px setzen
const powerPointTheme: Theme = {
  ...createLightTheme(powerPointBrand),
  // Base300 ist die Standard-Schriftgröße für fast alle Fluent UI Texte
  fontSizeBase300: "12px",
};

// 3. Custom Styles für die Tabs definieren
const useStyles = makeStyles({
  tabList: {
    width: '100%',
    columnGap: '4px',
  },
  customTab: {
    // 1. Globale Typografie für den Tab erzwingen
    fontSize: '11px',
    fontWeight: 'bold',

    // 2. Abstände oben/unten reduzieren (kompakterer Look)
    paddingTop: '2px',
    paddingBottom: '2px',
    paddingLeft: '12px',
    paddingRight: '12px',
    minHeight: '26px', // Vorher 32px
    borderRadius: tokens.borderRadiusMedium,

    // 3. Standard-Zustand (Unselected)
    color: tokens.colorNeutralForeground2,
    backgroundColor: 'transparent',

    // Versteckt den standardmäßigen Unterstrich
    '&::after': {
      display: 'none',
    },

    // 4. Hover-Zustand für UNSELECTED Tabs
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
      color: tokens.colorNeutralForeground2Hover, // Erzwingt, dass die Farbe beim Hovern dunkelgrau bleibt
    },

    // 5. Zustand für SELECTED Tabs
    '&[aria-selected="true"]': {
      backgroundColor: tokens.colorBrandBackground2,
      color: tokens.colorBrandForeground1,
    },

    // 6. Hover-Zustand für SELECTED Tabs (damit es beim Drüberfahren stimmig bleibt)
    '&[aria-selected="true"]:hover': {
      backgroundColor: tokens.colorBrandBackground2Hover,
      color: tokens.colorBrandForeground1,
    },
  }
});


export const App: React.FC = () => {
  const classes = useStyles();
  const [selectedValue, setSelectedValue] = useState<string>('templates');

  const onTabSelect = (_event: SelectTabEvent, data: SelectTabData) => {
    setSelectedValue(data.value as string);
  };

  return (
    <FluentProvider theme={powerPointTheme}>
      <div style={{ padding: '12px', display: 'flex', flexDirection: 'column', height: '100vh', boxSizing: 'border-box' }}>

        <TabList
          selectedValue={selectedValue}
          onTabSelect={onTabSelect}
          className={classes.tabList}
        >
          <Tab
            value="templates"
            className={classes.customTab}
          >
            Templates
          </Tab>

          <Tab
            value="icons"
            className={classes.customTab}
          >
            Symbole
          </Tab>

          <Tab
            value="settings"
            icon={<SettingsRegular />}
            aria-label="Settings"
            className={classes.customTab}
            style={{ marginLeft: 'auto' }}
          />
        </TabList>

        <div style={{ flex: 1, marginTop: '16px', overflow: 'hidden' }}>
          {selectedValue === 'templates' && <TemplatesTab />}
          {selectedValue === 'icons' && <IconsTab />}
          {selectedValue === 'settings' && <SettingsTab />}
        </div>
      </div>
    </FluentProvider>
  );
};