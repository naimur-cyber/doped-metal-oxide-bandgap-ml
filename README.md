# Bandgap Prediction in Doped Metal Oxides: Data, Code, and Manuscript

This repository contains the dataset, analysis code, figures, and manuscript files for a study on machine learning prediction of bandgap shifts (ΔEg) in doped metal-oxide semiconductors, with a particular focus on how inter-laboratory measurement variability limits model performance.

## Background

While compiling experimental ΔEg values from the literature for six oxide hosts (TiO₂, CeO₂, SnO₂, ZnO, In₂O₃, WO₃), it became clear that different research groups often report noticeably different bandgap values for what should be nominally identical doped systems. This study quantifies that inter-laboratory noise, compares it directly to the error of a trained ML model, and uses SHAP analysis to identify which features (and in particular, which synthesis parameters) drive the model's predictions and where its limitations come from.

## Dataset

- 875 unique data points collected from 200 sources (2005–2026); 638 remained after filtering and cleaning.
- 6 host oxides, 40 distinct dopants (transition metals, rare earths, and main-group elements).
- ΔEg ranges from -1.80 eV to +1.14 eV (mean 0.032 eV, std 0.446 eV), plus 126 undoped baseline measurements.
- Only data derived from Tauc-plot extrapolation of UV-Vis diffuse reflectance/transmission spectra was retained, to keep the bandgap determination method consistent across sources.

## Repository structure

```
├── data/
│   ├── raw/              # literature-extracted dataset before cleaning
│   └── processed/        # final 638-point dataset used for modeling
├── code/                  # scripts/notebooks for preprocessing, model training,
│                          # SHAP analysis, conformal prediction, LOHO validation
├── results/               # output tables (Tables 1-5, Table S1) and result sheets
├── figures/                # all figures used in the manuscript (Fig. 1-10)
└── manuscript/             # manuscript draft and supplementary material
```

(Adjust the above to match the actual folder names once everything is uploaded.)

## Summary of key results

The best-performing model (Gradient Boosting) achieved an MAE of 0.1699 eV on the full dataset, which is almost identical to the median inter-laboratory measurement variation (0.178 eV) computed independently from repeated measurements in the literature. After removing the most inconsistent data pairs, R² improved to 0.83 without any change to the model itself. Leave-One-Host-Out validation shows the model generalizes poorly across different oxide hosts and performs better as an interpolator within a host it has already seen. A within-paper analysis, where data were grouped by source to remove inter-lab noise, showed the model captures a real (if modest) concentration-dependent trend (r = 0.234, p = 0.048). SHAP analysis indicates that synthesis-related parameters are the most influential features overall.

The overall takeaway is that, for this problem, the model's accuracy is already close to the ceiling set by how consistently the underlying experiments were reported — so improving synthesis reporting standards in the literature is likely to matter more for future progress than further algorithmic tuning.

## Reproducing the results

1. Clone this repository.
2. Install dependencies: `pip install -r requirements.txt`
3. Run the preprocessing script to generate the cleaned dataset from `data/raw/`.
4. Run the modeling scripts in `code/` to reproduce the tables and figures in `results/` and `figures/`.

(Add exact commands/notebook names once finalized.)

## Citation

If you use this dataset or code, please cite:

> [Author names], "[Physics-Driven Insights into Inter-Laboratory Noise as the Limiting Factor in Machine Learning Prediction of Bandgap Shifts in Doped Metal Oxides]," [Journal], [year]. [DOI / link once available]

## License

[Specify license, e.g., MIT for code, CC-BY for data — update before making the repo public.]

## Contact

Questions or issues with the dataset or code can be raised via GitHub Issues, or contact [your email/contact info].
