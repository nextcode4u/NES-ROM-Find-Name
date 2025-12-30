# NES ROM Renamer

A PowerShell script that renames NES ROM files using DAT files.

Safe by default: it scans first and asks before renaming.

---

## What You Need

Put everything in **one folder**:

- `Rename NES by DAT.ps1`
- Your `.nes` ROM files
- One or more `.dat` files

---

## Recommended DAT

**OpenNES.dat**  
This gives the **best match rate** for mixed NES collections.

Optional (extra coverage):
- No-Intro (Headered)
- TOSEC

---

## Where to Get DAT Files

- OpenNES: https://github.com/SnowflakePowered/opengood community mirrors / GitHub
- No-Intro: https://datomatic.no-intro.org/
- TOSEC: https://www.tosecdev.org/

---

## How to Run

### Scan first (recommended)
```powershell
.\Rename NES by DAT.ps1
```

### Scan + clean 4-digit prefixes on unmatched files
```powershell
.\Rename NES by DAT.ps1 -Strip4DigitPrefixForUnmatched
```

### Rename automatically
```powershell
.\Rename NES by DAT.ps1 -Apply
```

---

## Notes on Performance

- The script **reads every ROM into memory** to calculate CRCs
- Large ROM sets or many DAT files can cause **high CPU and memory usage**
- This is normal and expected for CRC-based verification
- For best performance:
  - Close other heavy applications
  - Run in batches if you have very large collections

---

## Output Files

- `rename-log.txt`
- `planned-renames.csv`
- `unmatched.txt`

---

## License

MIT
