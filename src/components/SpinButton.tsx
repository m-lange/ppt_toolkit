import React, { useState, useEffect } from 'react';
import { Input, makeStyles, tokens } from '@fluentui/react-components';
import { ChevronUpRegular, ChevronDownRegular } from '@fluentui/react-icons';

const useStyles = makeStyles({
  container: {
    position: 'relative',
    width: '85px', // Standardbreite
  },
  // Container für die kleinen Pfeile
  spinControls: {
    display: 'flex',
    flexDirection: 'column',
    marginLeft: '2px',
  },
  // Die kleinen Pfeil-Buttons
  spinBtn: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    background: 'transparent',
    border: 'none',
    cursor: 'pointer',
    padding: '0px 4px',
    height: '12px',
    color: tokens.colorNeutralForeground3,
    '&:hover': {
      color: tokens.colorNeutralForeground1,
      backgroundColor: tokens.colorNeutralBackground1Hover,
      borderRadius: tokens.borderRadiusSmall,
    },
  },
  input: {
    width: '100%',
  }
});

interface SpinButtonProps {
  value: number;
  onChange: (newValue: number) => void;
  step?: number;
  min?: number;
  max?: number;
  suffix?: string;           // z.B. " cm"
  decimals?: number;         // z.B. 2
  decimalSeparator?: string; // z.B. "," für Deutsch oder "." für Englisch
}

export const SpinButton: React.FC<SpinButtonProps> = ({
  value,
  onChange,
  step = 1,
  min,
  max,
  suffix = "",
  decimals = 0,
  decimalSeparator = ","
}) => {
  const classes = useStyles();

  // Der Text, der tatsächlich im Eingabefeld steht (z.B. "1,50 cm")
  const [displayValue, setDisplayValue] = useState<string>("");

  // Formatiert eine echte Zahl in den gewünschten Text-String
  const formatValue = (val: number): string => {
    const numStr = val.toFixed(decimals).replace('.', decimalSeparator);
    return `${numStr}${suffix}`;
  };

  // Parst den Text-String zurück in eine echte Zahl
  const parseValue = (text: string): number => {
    let cleanText = text;
    if (suffix) cleanText = cleanText.replace(suffix, ''); // Suffix entfernen
    cleanText = cleanText.replace(decimalSeparator, '.');  // Komma zu Punkt machen
    cleanText = cleanText.replace(/[^0-9.-]/g, '');        // Alles außer Zahlen, Punkt und Minus entfernen

    const parsed = parseFloat(cleanText);
    return isNaN(parsed) ? 0 : parsed;
  };

  // Stellt sicher, dass die Zahl innerhalb von min/max bleibt
  const clamp = (val: number): number => {
    let res = val;
    if (min !== undefined && res < min) res = min;
    if (max !== undefined && res > max) res = max;
    return res;
  };

  // Wenn sich der Wert von außen ändert, aktualisieren wir die Anzeige
  useEffect(() => {
    setDisplayValue(formatValue(value));
  }, [value, decimals, decimalSeparator, suffix]);

  // Wird aufgerufen, wenn der Nutzer tippt (lässt freies Tippen zu)
  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setDisplayValue(e.target.value);
  };

  // Wird aufgerufen, wenn das Feld verlassen wird (Speichern)
  const handleBlur = () => {
    const parsed = parseValue(displayValue);
    const clamped = clamp(parsed);
    onChange(clamped);
    setDisplayValue(formatValue(clamped)); // Formatiert es wieder sauber (z.B. hängt " cm" wieder an)
  };

  // Klick auf Pfeil Hoch
  const handleIncrement = () => {
    const current = parseValue(displayValue);
    const clamped = clamp(current + step);
    onChange(clamped);
    setDisplayValue(formatValue(clamped));
  };

  // Klick auf Pfeil Runter
  const handleDecrement = () => {
    const current = parseValue(displayValue);
    const clamped = clamp(current - step);
    onChange(clamped);
    setDisplayValue(formatValue(clamped));
  };

  // Erlaubt die Steuerung mit den Pfeiltasten der Tastatur!
  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'ArrowUp') {
      e.preventDefault();
      handleIncrement();
    } else if (e.key === 'ArrowDown') {
      e.preventDefault();
      handleDecrement();
    } else if (e.key === 'Enter') {
      e.preventDefault();
      handleBlur(); // Speichert bei Enter
    }
  };

  return (
    <div className={classes.container}>
      <Input
        className={classes.input}
        value={displayValue}
        onChange={handleChange}
        onBlur={handleBlur}
        onKeyDown={handleKeyDown}
        appearance="filled-darker"
        size="small"
        // Hier fügen wir unsere eigenen Pfeile rechts ins Eingabefeld ein
        contentAfter={
          <div className={classes.spinControls}>
            <button type="button" tabIndex={-1} className={classes.spinBtn} onClick={handleIncrement}>
              <ChevronUpRegular fontSize={10} />
            </button>
            <button type="button" tabIndex={-1} className={classes.spinBtn} onClick={handleDecrement}>
              <ChevronDownRegular fontSize={10} />
            </button>
          </div>
        }
      />
    </div>
  );
};