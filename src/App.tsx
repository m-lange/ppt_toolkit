import React, { useState } from 'react';
import {
  FluentProvider,
  webLightTheme,
  TabList,
  Tab,
  type SelectTabData,
  type SelectTabEvent
} from '@fluentui/react-components';

import { TemplatesTab } from './components/TemplatesTab';
import { SettingsTab } from './components/SettingsTab';

interface AppProps {
  isOffice?: boolean;
}

export const App: React.FC<AppProps> = ({ isOffice }) => {
  const [selectedValue, setSelectedValue] = useState<string>('templates');

  // Typisierung für das Fluent UI Tab-Event
  const onTabSelect = (event: SelectTabEvent, data: SelectTabData) => {
    setSelectedValue(data.value as string);
  };

  return (
    <FluentProvider theme={webLightTheme}>
      <div style={{ padding: '16px', display: 'flex', flexDirection: 'column', height: '100vh', boxSizing: 'border-box' }}>
        <TabList selectedValue={selectedValue} onTabSelect={onTabSelect} size="medium">
          <Tab value="templates">Templates</Tab>
          <Tab value="settings">Settings</Tab>
        </TabList>

        <div style={{ flex: 1, marginTop: '24px' }}>
          {selectedValue === 'templates' && <TemplatesTab />}
          {selectedValue === 'settings' && <SettingsTab />}
        </div>
      </div>
    </FluentProvider>
  );
};