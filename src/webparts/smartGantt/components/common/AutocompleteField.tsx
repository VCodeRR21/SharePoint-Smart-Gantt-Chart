import * as React from 'react';
import { Label } from '@fluentui/react';
import styles from './AutocompleteField.module.scss';

interface IAutocompleteFieldProps {
  label?: string;
  value: string;
  suggestions: string[];
  onChange: (value: string) => void;
  placeholder?: string;
  required?: boolean;
}

export const AutocompleteField: React.FC<IAutocompleteFieldProps> = ({
  label, value, suggestions, onChange, placeholder, required,
}) => {
  const [open, setOpen] = React.useState(false);
  const [highlighted, setHighlighted] = React.useState(-1);
  const wrapRef = React.useRef<HTMLDivElement>(null);

  const filtered = value
    ? suggestions.filter(s => s.toLowerCase().includes(value.toLowerCase()) && s !== value)
    : suggestions;
  // Only the first 12 matches are ever rendered — clamp keyboard navigation
  // and Enter-to-commit to that same window, or ArrowDown past the 12th item
  // could highlight (and Enter could commit) an option the user never saw.
  const visible = filtered.slice(0, 12);

  // Close on outside click
  React.useEffect(() => {
    const handler = (e: MouseEvent): void => {
      if (wrapRef.current && !wrapRef.current.contains(e.target as Node)) {
        setOpen(false);
        setHighlighted(-1);
      }
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, []);

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>): void => {
    if (!open || visible.length === 0) {
      if (e.key === 'ArrowDown' && visible.length > 0) { setOpen(true); setHighlighted(0); }
      return;
    }
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      setHighlighted(h => Math.min(h + 1, visible.length - 1));
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      setHighlighted(h => Math.max(h - 1, 0));
    } else if (e.key === 'Enter' && highlighted >= 0) {
      e.preventDefault();
      onChange(visible[highlighted]);
      setOpen(false);
      setHighlighted(-1);
    } else if (e.key === 'Escape') {
      setOpen(false);
      setHighlighted(-1);
    }
  };

  return (
    <div className={styles.wrap} ref={wrapRef}>
      {label && <Label required={required}>{label}</Label>}
      <input
        className={styles.input}
        value={value}
        placeholder={placeholder}
        onChange={e => { onChange(e.target.value); setOpen(true); setHighlighted(-1); }}
        onFocus={() => setOpen(true)}
        onKeyDown={handleKeyDown}
        autoComplete="off"
      />
      {open && visible.length > 0 && (
        <div className={styles.dropdown}>
          {visible.map((s, i) => (
            <div
              key={s}
              className={`${styles.option} ${i === highlighted ? styles.highlighted : ''}`}
              onMouseDown={e => { e.preventDefault(); onChange(s); setOpen(false); setHighlighted(-1); }}
            >
              {s}
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

export default AutocompleteField;
