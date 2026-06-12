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


const brandTheme: BrandVariants = {
  10: "#020401",
  20: "#101C0F",
  30: "#152F17",
  40: "#173D1C",
  50: "#194B20",
  60: "#1A5A25",
  70: "#1A692A",
  80: "#19782E",
  90: "#168833",
  100: "#119938",
  110: "#07A93C",
  120: "#34B84E",
  130: "#60C46B",
  140: "#82D087",
  150: "#A2DCA3",
  160: "#BFE7BF"
};

// 2. Theme anpassen: Globale Schriftgröße auf 12px setzen
const powerPointTheme: Theme = {
  ...createLightTheme(brandTheme),
  fontSizeBase300: "12px",
};

// 3. Custom Styles für die Tabs definieren
const useStyles = makeStyles({
  tabList: {
    width: '100%',
    columnGap: '4px',
  },
  customTab: {
    fontSize: '11px',
    fontWeight: 'bold',
    paddingTop: '2px',
    paddingBottom: '2px',
    paddingLeft: '12px',
    paddingRight: '12px',
    minHeight: '26px',
    borderRadius: tokens.borderRadiusMedium,

    color: tokens.colorNeutralForeground2,
    backgroundColor: 'transparent',

    '&::after': {
      display: 'none !important',
    },

    '&:hover': {
      color: tokens.colorNeutralForeground1,
      backgroundColor: 'transparent',
    },

    '&[aria-selected="true"]': {
      backgroundColor: tokens.colorBrandBackground2,
      color: tokens.colorBrandForeground1,
    },

    '&[aria-selected="true"]:hover': {
      backgroundColor: tokens.colorBrandBackground2,
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