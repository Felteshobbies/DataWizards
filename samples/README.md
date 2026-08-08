# Sample files

Each file exercises a case that is easy to get wrong. Drop them onto the Convert
tab and read the analysis panel.

| File | What it demonstrates |
|---|---|
| `leading-zeros-and-german-numbers.csv` | `00123` and `01067` stay text; `19,99` becomes 19.99; `31.01.2025` becomes a real date |
| `semicolon-quoted.csv` | quoted values containing `\|` and `,` do not split the record |
| `pipe-separated.csv` | pipe separator, with a `\|` inside a quoted value |
| `tab-separated-no-header.csv` | tab separator, and a file with no header row |
| `windows-1252-umlauts.csv` | a legacy single-byte encoding. Detection guesses ibm850 on a file this short and flags the low confidence; set the encoding to `windows-1252` on the Detection tab to read it correctly |
| `with-preamble-lines.csv` | two comment lines before the real header. Set *skip leading lines* to 2 on the Detection tab; without it the analysis warns that the field count varies |
