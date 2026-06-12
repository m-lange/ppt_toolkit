import React, { useState, useEffect } from 'react';
import {
  SearchBox,
  Accordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
  makeStyles,
  tokens,
  Tooltip
} from '@fluentui/react-components';
import { type CategoryConfig, type MainIconConfig } from '../types/iconConfig';
import { tsIconCategories } from '../config/icons';

// Styling für das Icon-Grid und die flachen Buttons
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
    backgroundColor: tokens.colorNeutralBackground3, // Leicht grauer Hintergrund
    color: tokens.colorNeutralForeground1,
    cursor: 'pointer',
    transition: 'background-color 0.2s',
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground3Hover, // Dunkler beim Hover
    },
    '& svg': {
      width: '32px',
      height: '32px',
    },
    '& img': {
      width: '32px',
      height: '32px',
    }
  }
});

export const IconsTab: React.FC = () => {
  const classes = useStyles();
  const [allCategories, setAllCategories] = useState<CategoryConfig[]>([]);
  const [searchText, setSearchText] = useState<string>('');

  // Daten beim ersten Rendern laden und zusammenführen
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

        // TS-Daten und JSON-Daten kombinieren
        setAllCategories([...tsIconCategories, ...jsonCategories]);
      } catch (error) {
        console.error("Fehler beim Laden der JSON-Icons:", error);
        // Falls JSON fehlschlägt, zumindest TS-Icons anzeigen
        setAllCategories(tsIconCategories);
      }
    };

    loadData();
  }, []);

  // Filter-Logik
  const filteredCategories = allCategories.map(category => {
    const term = searchText.toLowerCase();
    const filteredIcons = category.icons.filter(icon => {
      const matchName = icon.name.toLowerCase().includes(term);
      const matchKeyword = icon.keywords.some(kw => kw.toLowerCase().includes(term));
      return matchName || matchKeyword;
    });

    return { ...category, icons: filteredIcons };
  }).filter(category => category.icons.length > 0); // Leere Kategorien ausblenden

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
      {/* Underline Searchbar */}
      <SearchBox
        appearance="underline"
        placeholder="Icons suchen..."
        onChange={(_e, data) => setSearchText(data.value || '')}
      />

      {/* Accordion für die Kategorien */}
      <div style={{ overflowY: 'auto', flex: 1, paddingRight: '4px' }}>
        {filteredCategories.length === 0 ? (
          <div style={{ textAlign: 'center', marginTop: '20px', color: tokens.colorNeutralForeground3 }}>
            Keine Icons gefunden.
          </div>
        ) : (
          <Accordion multiple collapsible defaultOpenItems={filteredCategories.map(c => c.catagory)}>
            {filteredCategories.map((category, index) => (
              <AccordionItem value={category.category} key={index}>
                <AccordionHeader>{category.category}</AccordionHeader>
                <AccordionPanel>
                  <div className={classes.grid}>
                    {category.icons.map((icon, i) => (
                      <Tooltip content={icon.name} relationship="label" key={i}>
                        <button
                          className={classes.iconButton}
                          onClick={() => console.log(`Icon ${icon.name} geklickt!`)}
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