import React, { useState, useEffect } from 'react';
import { SearchBox, Button, Field } from '@fluentui/react-components';
import { ShapesRegular, MaximizeRegular, CodeRegular } from '@fluentui/react-icons';
import { insertBlueRectangle, showSlideDimensions, exportSelectedShapeToJson } from '../utils/powerpointApi';

export const TemplatesTab: React.FC = () => {
  const [isShapeSelected, setIsShapeSelected] = useState(false);

  // Lauscht auf Selektionsänderungen in PowerPoint
  useEffect(() => {
    const checkSelection = async () => {
      if (typeof PowerPoint === 'undefined') return;
      try {
        await PowerPoint.run(async (context) => {
          const shapes = context.presentation.getSelectedShapes();
          shapes.load("items");
          await context.sync();
          // Wenn mindestens 1 Element selektiert ist, wird der Button aktiv
          setIsShapeSelected(shapes.items.length > 0);
        });
      } catch (error) {
        // "ItemNotFound" wird geworfen, wenn man ins Leere klickt
        setIsShapeSelected(false);
      }
    };

    const handleSelectionChange = () => {
      checkSelection();
    };

    if (typeof Office !== 'undefined' && Office.context && Office.context.document) {
      // Event-Listener registrieren
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        handleSelectionChange
      );
      // Initialer Check beim Öffnen des Tabs
      checkSelection();
    }

    return () => {
      // Event-Listener beim Verlassen des Tabs sauber entfernen
      if (typeof Office !== 'undefined' && Office.context && Office.context.document) {
        Office.context.document.removeHandlerAsync(
          Office.EventType.DocumentSelectionChanged,
          { handler: handleSelectionChange }
        );
      }
    };
  }, []);

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
      <Field label="Vorlagen durchsuchen">
        <SearchBox placeholder="Suche..." size="small" />
      </Field>

      <Field label="Verfügbare Vorlagen">
        <Button
          appearance="primary"
          icon={<ShapesRegular />}
          onClick={insertBlueRectangle}
          size="small"
        >
          Blaues Rechteck einfügen
        </Button>
      </Field>

      {/* NEU: Button 1 - Foliengröße */}
      <Field label="Folieneigenschaften">
        <Button
          appearance="secondary"
          icon={<MaximizeRegular />}
          onClick={showSlideDimensions}
          size="small"
        >
          Foliengröße auslesen
        </Button>
      </Field>

      {/* NEU: Button 2 - JSON Export */}
      <Field
        label="Objekt-Export"
        validationMessage={!isShapeSelected ? "Bitte wähle zuerst ein Objekt auf der Folie aus." : "Objekt ausgewählt. Bereit zum Export."}
      >
        <Button
          appearance="secondary"
          icon={<CodeRegular />}
          onClick={exportSelectedShapeToJson}
          size="small"
          disabled={!isShapeSelected} // Button ist ausgegraut, wenn nichts selektiert ist
        >
          Eigenschaften exportieren (JSON)
        </Button>
      </Field>
    </div>
  );
};