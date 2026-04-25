# AutoChi

(https://autochi.streamlit.app)[https://autochi.streamlit.app]

AutoChi is a Streamlit-based statistical analysis app for automated chi-square analysis of proportions with one or two categorical factors and two response-count columns.

## Expected input structure

The app assumes the uploaded CSV/XLSX file has this layout:

- First **1 or 2 columns**: categorical factor column(s)
- Last **2 columns**: response-count columns
- First row: variable names

Examples:

### One-factor layout
| Country | Yes | No |
|---------|-----|----|
| UAE     | 120 | 80 |
| Oman    | 90  | 110 |

### Two-factor layout
| Country | Platform | Yes | No |
|---------|----------|-----|----|
| UAE     | Facebook | 120 | 80 |
| UAE     | Instagram| 140 | 60 |

## Main features

- Upload CSV or XLSX files
- Automatic handling of **1-factor** and **2-factor** designs
- Chi-square test from contingency tables
- Pairwise two-proportion z-tests
- Bonferroni correction
- Directed significance network visualization
- Arc diagram visualization
- Export results to Word (`.docx`)

## Required files for deployment

This repository should contain:

- `app.py`
- `requirements.txt`
- `packages.txt`
- `.streamlit/config.toml`

## Local run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud

1. Push this repository to GitHub.
2. Sign in to Streamlit Community Cloud with GitHub.
3. Click **Create app**.
4. Select your repository, branch, and `app.py`.
5. Deploy.

## Notes

- `graphviz` is included in both `requirements.txt` and `packages.txt` because the Python package and the system binary are both needed for full network rendering.
- If Graphviz is not available, the app still runs, but network PNG export may be unavailable.
