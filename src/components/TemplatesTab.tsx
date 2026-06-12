import React from 'react';
import { Button, Field } from '@fluentui/react-components';
import { ShapesRegular } from '@fluentui/react-icons';
import { insertBlueRectangle } from '../utils/powerpointApi';

export const TemplatesTab: React.FC = () => {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>

      <Field label="Verfügbare Vorlagen" validationMessage="Klicke auf den Button, um eine Form in die aktive Folie einzufügen.">
        <Button
          appearance="primary"
          icon={<ShapesRegular />}
          onClick={insertBlueRectangle}
        >
          Blaues Rechteck einfügen
        </Button>
      </Field>
    </div>
  );
};