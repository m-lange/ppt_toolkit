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

const useStyles = makeStyles({
  grid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fill, minmax(60px, 1fr))',
    gap: '8px',
    padding: '4px 0',
  },
  iconButton: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    height: '60px',
    border: 'none',
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground1,
    cursor: 'pointer',
    transition: 'background-color 0.2s',
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground3Hover,
    },
    '& svg': {
      width: '32px',
      height: '32px',
    },
    '& img': {
      width: '32px',
      height: '32px',
    }
  },
  searchBox: {
    backgroundColor: tokens.colorNeutralBackground3, // Gleiches Grau wie die Buttons
    borderRadius: tokens.borderRadiusMedium, // Runde Ecken

    // Optional: Versteckt den farbigen Strich am unteren Rand, wenn man reinklickt
    '&::after': {
      display: 'none',
    }
  },

  accordionHeader: {
    margin: '0px',
    '& button': {
      minHeight: '32px', // Reduziert die Höhe des Headers
      paddingLeft: '4px', // Reduziert den Abstand links
      paddingTop: '0px',
      paddingBottom: '0px',
    }
  },

  // NEU: Kompaktere Abstände für den Inhalt (Panel)
  accordionPanel: {
    margin: '0px',
    paddingTop: '4px',
    paddingBottom: '8px', // Leichter Abstand nach unten zur nächsten Kategorie
    paddingLeft: '4px', // Reduziert den Abstand links, damit es bündig mit dem Header ist
  }
});

export const IconsTab: React.FC = () => {
  const classes = useStyles();
  const [allCategories, setAllCategories] = useState<CategoryConfig[]>([]);
  const [searchText, setSearchText] = useState<string>('');

  // NEU: Gesteuerter Zustand für das Accordion (Startet als leeres Array = alles geschlossen)
  const [openItems, setOpenItems] = useState<string[]>([]);

  useEffect(() => {
    const loadData = async () => {
      try {
        const res = await fetch(`${import.meta.env.BASE_URL}config/icons.json`);
        const mainConfig: MainIconConfig = await res.json();

        const jsonPromises = mainConfig.files.map(async (file) => {
          const cleanPath = file.replace('./', import.meta.env.BASE_URL);
          const catRes = await fetch(cleanPath);
          return catRes.json() as Promise<CategoryConfig>;
        });

        const jsonCategories = await Promise.all(jsonPromises);
        setAllCategories([...tsIconCategories, ...jsonCategories]);
      } catch (error) {
        console.error("Fehler beim Laden der JSON-Icons:", error);
        setAllCategories(tsIconCategories);
      }
    };

    loadData();
  }, []);

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

  // NEU: Such-Handler (Öffnet Kategorien automatisch)
  const handleSearchChange = (_event: SearchBoxChangeEvent, data: { value?: string }) => {
    const text = data.value || '';
    setSearchText(text);

    if (text.trim() !== '') {
      // Wenn gesucht wird: Finde alle Kategorien, die Treffer haben, und öffne sie
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
      // Wenn die Suche geleert wird: Schließe alle Accordions wieder
      setOpenItems([]);
    }
  };

  // NEU: Manueller Klick auf einen Accordion-Header
  const handleToggle = (_event: AccordionToggleEvent, data: AccordionToggleData) => {
    setOpenItems(data.openItems as string[]);
  };

  // NEU: Klick auf ein Icon (Einfügen in PowerPoint)
  const handleIconClick = async (svgData: string | string[]) => {
    const svgString = Array.isArray(svgData) ? svgData.join('\n') : svgData;

    if (svgString.trim().startsWith('<svg')) {
      // Inline SVG direkt einfügen
      await insertSvgIcon(svgString);
    } else {
      // Falls es eine URL ist, laden wir den SVG-Code erst herunter
      try {
        const res = await fetch(svgString);
        const text = await res.text();
        await insertSvgIcon(text);
      } catch (err) {
        console.error("Fehler beim Laden des SVGs von URL:", err);
      }
    }
  };

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
        appearance="filled-darker"
        placeholder="Symbole suchen..."
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
                {/* NEU: <b> Tag macht den Header fett */}
                <AccordionHeader  className={classes.accordionHeader}>
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