# Final Test Analyzer

This Python project analyzes calibration test data collected from sensor mats. It detects the locations of applied weights, calculates the differential and absolute capacitance values at each region, classifies the pressure levels using color mapping, and generates a comprehensive PowerPoint report summarizing the results.

## 🔍 Features

- ✅ Automatic detection of weight spots using differential matrix analysis
- 📈 Calculation of average and maximum capacitance values per weight (5lb, 10lb, 20lb)
- 🎨 Pressure level classification into color zones (Red, Orange, Yellow, Green, etc.)
- 📊 Visual and tabular summary of results across multiple sensor mats
- 📤 Generation of:
  - PowerPoint report (`.pptx`)
  - Color distribution pie charts
  - Summary tables
- 🐛 Optional debug mode for interactive visualization

## 📁 Folder Structure

The input should be organized as follows:

```
<base_folder>/
│
├── Mat3000/
│   ├── 3000_rawData.csv
│   ├── 3000_withWeight_rawData.csv
│   └── 3000_withWeight_heatmap.png
│
├── Mat3001/
│   ├── ...
│
└── ...
```

Each subfolder should contain:
- Two CSV files (one before weights, one after)
- Optional: heatmap image (`*_heatmap.png`) for inclusion in the slides

## 🧪 CSV Format

- Each CSV file contains **3 matrices**, separated by lines that start with `row_0`
- The **third matrix** contains the capacitance readings and is used for analysis
- The **second matrix** is used to determine the pressure level color for each weight area

## ▶️ Usage

### Run from command line:

```bash
python main.py
```

> 💡 Make sure to update the `base_folder` variable in `main.py` before running.

Or modify the main block:
```python
if __name__ == "__main__":
    base_folder = r"path\to\your\mat\folders"
    analyze_all(base_folder, debug=True)  # Enable debug visualization
```

## 📦 Requirements

Install dependencies with pip:

```bash
pip install pandas numpy matplotlib scipy python-pptx
```

## 📤 Output

After running the analysis, the following will be generated in the base folder:

- `final_test_analysis_report.pptx` — A PowerPoint presentation with:
  - One slide per mat with weight stats and color
  - Summary table slides (split into pages)
  - Pie charts showing color distribution per weight
- (optional) Add Excel summary if you include `save_excel_with_summary()` function

## 📌 Notes

- The script auto-selects the "empty" vs. "with weight" CSV based on average values
- It assumes weights are placed in roughly fixed vertical positions (sorted by Y)

## 🛠 Future Improvements

- Export Excel summary with charts and comparisons
- GUI for selecting folders and thresholds
- Improved weight area detection using ML or spatial heuristics

---

© 2025 – Developed by Dvir Shavit for Wellsense R&D
