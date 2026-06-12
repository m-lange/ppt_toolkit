import React, { useState } from 'react';
import {
  FluentProvider,
  createLightTheme,
  TabList,
  Tab,
  type BrandVariants,
  type Theme,
  type SelectTabData,
  type SelectTabEvent
} from '@fluentui/react-components';
import { TemplatesTab } from './components/TemplatesTab';
import { SettingsTab } from './components/SettingsTab';


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
  100: "#C43E1C",
  110: "#D05232",
  120: "#DB6547",
  130: "#E4785D",
  140: "#EC8B73",
  150: "#F39E8A",
  160: "#F8B1A1",
};


const powerPointTheme: Theme = {
  ...createLightTheme(powerPointBrand),
  fontSizeBase100: "8px",
  fontSizeBase200: "10px",
  fontSizeBase300: "12px",
  fontSizeBase400: "14px",
  fontSizeBase500: "16px",
  fontSizeBase600: "18px",
};

interface AppProps {
  isOffice?: boolean;
}

export const App: React.FC<AppProps> = ({ isOffice }) => {
  const [selectedValue, setSelectedValue] = useState<string>('templates');

  const onTabSelect = (event: SelectTabEvent, data: SelectTabData) => {
    setSelectedValue(data.value as string);
  };

  return (
    <FluentProvider theme={powerPointTheme}>
      <div style={{ padding: '12px', display: 'flex', flexDirection: 'column', height: '100vh', boxSizing: 'border-box' }}>

        {/* size="small" macht die Tabs kompakter */}
        <TabList selectedValue={selectedValue} onTabSelect={onTabSelect} size="small">
          <Tab value="templates">Templates</Tab>
          <Tab value="settings">Settings</Tab>
        </TabList>

        <div style={{ flex: 1, marginTop: '16px' }}>
          {selectedValue === 'templates' && <TemplatesTab />}
          {selectedValue === 'settings' && <SettingsTab />}
        </div>
      </div>
    </FluentProvider>
  );
};