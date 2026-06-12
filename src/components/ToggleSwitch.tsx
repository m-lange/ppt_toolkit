import React from 'react';
import { Button, makeStyles, tokens } from '@fluentui/react-components';
import { CheckmarkRegular } from '@fluentui/react-icons';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'flex-start',
  },
  toggleButton: {
    width: '24px',  // Exakt die Höhe eines "small" Input-Feldes
    height: '24px',
    minWidth: '24px',
    padding: '0px',

    // Formatierung analog zu appearance="filled-darker"
    backgroundColor: tokens.colorNeutralBackground3,
    border: 'none',
    borderRadius: tokens.borderRadiusMedium,
    color: tokens.colorNeutralForeground1,

    // Hover-Effekt analog zum Input-Feld
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground3Hover,
      color: tokens.colorNeutralForeground1Hover,
    },

    // Klick-Effekt
    '&:active': {
      backgroundColor: tokens.colorNeutralBackground3Pressed,
      color: tokens.colorNeutralForeground1Pressed,
    },

    // Macht das Icon minimal kleiner, damit es perfekt in die 24x24px Box passt
    '& svg': {
      width: '16px',
      height: '16px',
    }
  }
});

interface ToggleSwitchProps {
  checked: boolean;
  onChange: (checked: boolean) => void;
  width?: string; // Standardmäßig 85px, damit es bündig mit den SpinButtons bleibt
}

export const ToggleSwitch: React.FC<ToggleSwitchProps> = ({
  checked,
  onChange,
  width = '85px'
}) => {
  const classes = useStyles();

  return (
    <div className={classes.container} style={{ width }}>
      <Button
        className={classes.toggleButton}
        appearance="subtle" // Entfernt alle Standard-Rahmen von Microsoft
        onClick={() => onChange(!checked)}
        // Wenn checked true ist, zeigen wir den Haken.
        // Wenn false, zeigen wir ein unsichtbares 16x16 Element, damit der Button nicht schrumpft!
        icon={checked ? <CheckmarkRegular /> : <div style={{ width: '16px', height: '16px' }} />}
      />
    </div>
  );
};