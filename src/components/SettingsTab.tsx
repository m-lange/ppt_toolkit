import React from 'react';
import { Button, Text, Field } from '@fluentui/react-components';
import { SettingsRegular } from '@fluentui/react-icons';
import { insertGreenCircle } from '../utils/powerpointApi';

export const SettingsTab: React.FC = () => {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
      <Text>Hier könnten in Zukunft Add-In Einstellungen vorgenommen werden.</Text>

      <Field label="Test Aktionen">
        <Button
          appearance="secondary"
          icon={<SettingsRegular />}
          onClick={insertGreenCircle}
          size="small"
        >
          Grünen Kreis einfügen
        </Button>
      </Field>
    </div>
  );
};