import React, { useState, useEffect } from 'react';
import {
  SearchBox,
  Accordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
  makeStyles,
  tokens,
  Tooltip,
  type AccordionToggleEvent,
  type AccordionToggleData,
  type SearchBoxChangeEvent
} from '@fluentui/react-components';
import { type CategoryConfig, type MainIconConfig } from '../types/iconConfig';
import { tsIconCategories } from '../config/icons';
import { insertSvgIcon } from '../utils/powerpointApi';

import { loadSetting } from '../utils/settings';

const useStyles = makeStyles({
  grid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fill, minmax(48px, 1fr))', // Kompaktere Spalten
    gap: '8px',
    padding: '4px 0',
  },
  iconButton: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    height: '48px', // Kompaktere Button-Höhe
    border: 'none',
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground1,
    cursor: 'pointer',
    transition: 'background-color 0.2s',
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground3Hover,
    },
    '& svg': { width: '24px', height: '24px' }, // Kleinere Icons
    '& img': { width: '24px', height: '24px' }
  },
  searchBox: {
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
    '&::after': { display: 'none' }
  },
  accordionHeader: {
    margin: '0px',
    '& button': {
      minHeight: '32px',
      paddingLeft: '4px',
      paddingTop: '0px',
      paddingBottom: '0px',
    }
  },
  accordionPanel: {
    margin: '0px',
    paddingTop: '4px',
    paddingBottom: '8px',
    paddingLeft: '4px',
  }
});

interface IconsTabProps {
  reloadTrigger: number;
}

export const IconsTab: React.FC<IconsTabProps> = ({ reloadTrigger }) => {
  const classes = useStyles();
  const [allCategories, setAllCategories] = useState<CategoryConfig[]>([]);
  const [searchText, setSearchText] = useState<string>('');
  const [openItems, setOpenItems] = useState<string[]>([]);

  useEffect(() => {
    const loadData = async () => {
      try {
        // 1. Standard-URL definieren
        let configUrl = `${import.meta.env.BASE_URL}config/icons.json`;

        const savedUrl = loadSetting('iconUrl', '', false);
        if (savedUrl && savedUrl.trim() !== '') {
          configUrl = savedUrl;
        }

        // 3. Daten von der URL laden
        const res = await fetch(configUrl);
        const mainConfig: MainIconConfig = await res.json();

        const jsonPromises = mainConfig.files.map(async (file) => {
          // Falls die eigene URL genutzt wird, müssen relative Pfade korrekt aufgelöst werden
          const isAbsolute = file.startsWith('http://') || file.startsWith('https://');
          const cleanPath = isAbsolute
            ? file
            : new URL(file, configUrl.substring(0, configUrl.lastIndexOf('/') + 1)).href;

          const catRes = await fetch(cleanPath);
          return catRes.json() as Promise<CategoryConfig>;
        });

        const jsonCategories = await Promise.all(jsonPromises);
        setAllCategories([...tsIconCategories, ...jsonCategories]);

      } catch (error) {
        console.error("Fehler beim Laden der Icons (URL evtl. ungültig):", error);
        // Fallback: Nur die fest einkompilierten TS-Icons anzeigen
        setAllCategories(tsIconCategories);
      }
    };

    loadData();
  }, [reloadTrigger]); // Lädt neu, wenn der Trigger aus den Settings ausgelöst wird

  // Filter-Logik
  const filteredCategories = allCategories.map(cat => {
    const term = searchText.toLowerCase();
    const filteredIcons = cat.icons.filter(icon => {
      const matchName = icon.name.toLowerCase().includes(term);
      const matchKeyword = icon.keywords.some(kw => kw.toLowerCase().includes(term));
      return matchName || matchKeyword;
    });
    return { ...cat, icons: filteredIcons };
  }).filter(cat => cat.icons.length > 0);

  // Such-Handler (Öffnet Kategorien automatisch)
  const handleSearchChange = (_event: SearchBoxChangeEvent, data: { value?: string }) => {
    const text = data.value || '';
    setSearchText(text);

    if (text.trim() !== '') {
      const matchingCategories = allCategories
        .filter(cat => {
          const term = text.toLowerCase();
          return cat.icons.some(icon =>
            icon.name.toLowerCase().includes(term) ||
            icon.keywords.some(kw => kw.toLowerCase().includes(term))
          );
        })
        .map(cat => cat.category);
      setOpenItems(matchingCategories);
    } else {
      setOpenItems([]);
    }
  };

  // Manueller Klick auf einen Accordion-Header
  const handleToggle = (_event: AccordionToggleEvent, data: AccordionToggleData) => {
    setOpenItems(data.openItems as string[]);
  };

  // Klick auf ein Icon (Einfügen in PowerPoint)
  const handleIconClick = async (svgData: string | string[]) => {
    const svgString = Array.isArray(svgData) ? svgData.join('\n') : svgData;

    if (svgString.trim().startsWith('<svg')) {
      await insertSvgIcon(svgString);
    } else {
      try {
        const res = await fetch(svgString);
        const text = await res.text();
        await insertSvgIcon(text);
      } catch (err) {
        console.error("Fehler beim Laden des SVGs von URL:", err);
      }
    }
  };

  // Hilfsfunktion zum Rendern des SVGs (Inline oder URL)
  const renderSvg = (svgData: string | string[]) => {
    const svgString = Array.isArray(svgData) ? svgData.join('\n') : svgData;
    if (svgString.trim().startsWith('<svg')) {
      return <div dangerouslySetInnerHTML={{ __html: svgString }} style={{ display: 'flex' }} />;
    } else {
      return <img src={svgString} alt="icon" />;
    }
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: '16px', height: '100%' }}>
      <SearchBox
        className={classes.searchBox}
        appearance="filled-darker"
        placeholder="Icons suchen..."
        onChange={handleSearchChange}
      />

      <div style={{ overflowY: 'auto', flex: 1, paddingRight: '4px' }}>
        {filteredCategories.length === 0 ? (
          <div style={{ textAlign: 'center', marginTop: '20px', color: tokens.colorNeutralForeground3 }}>
            Keine Icons gefunden.
          </div>
        ) : (
          <Accordion
            multiple
            collapsible
            openItems={openItems}
            onToggle={handleToggle}
          >
            {filteredCategories.map((cat, index) => (
              <AccordionItem value={cat.category} key={index}>
                <AccordionHeader className={classes.accordionHeader}>
                  <b>{cat.category}</b>
                </AccordionHeader>
                <AccordionPanel className={classes.accordionPanel}>
                  <div className={classes.grid}>
                    {cat.icons.map((icon, i) => (
                      <Tooltip content={icon.name} relationship="label" key={i}>
                        <button
                          className={classes.iconButton}
                          onClick={() => handleIconClick(icon.svg)}
                        >
                          {renderSvg(icon.svg)}
                        </button>
                      </Tooltip>
                    ))}
                  </div>
                </AccordionPanel>
              </AccordionItem>
            ))}
          </Accordion>
        )}
      </div>
    </div>
  );
};