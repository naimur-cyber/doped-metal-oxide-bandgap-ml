Bandgap Prediction in Doped Metal Oxides: Data, Code, and Manuscript

This repository contains the dataset, analysis code, figures, and manuscript files for a study on machine learning prediction of bandgap shifts (ΔEg) in doped metal-oxide semiconductors, with a particular focus on how inter-laboratory measurement variability limits model performance.

Background

While compiling experimental ΔEg values from the literature for six oxide hosts (TiO₂, CeO₂, SnO₂, ZnO, In₂O₃, WO₃), it became clear that different research groups often report noticeably different bandgap values for what should be nominally identical doped systems. This study quantifies that inter-laboratory noise, compares it directly to the error of a trained ML model, and uses SHAP analysis to identify which features (and in particular, which synthesis parameters) drive the model's predictions and where its limitations come from.

Dataset


875 unique data points collected from 200 sources (2005–2026); 638 remained after filtering and cleaning.
6 host oxides, 40 distinct dopants (transition metals, rare earths, and main-group elements).
ΔEg ranges from -1.80 eV to +1.14 eV (mean 0.032 eV, std 0.446 eV), plus 126 undoped baseline measurements.
Only data derived from Tauc-plot extrapolation of UV-Vis diffuse reflectance/transmission spectra was retained, to keep the bandgap determination method consistent across sources.


Repository structure

├── bandgap_outputs/                        # output figures and result tables
│   ├── fig01a_r2_comparison.tiff
│   ├── fig01b_rmse.tiff
│   ├── fig01c_mae.tiff
│   ├── fig02a_parity_extra_trees.tiff
│   ├── fig02b_parity_random_forest.tiff
│   ├── fig02c_parity_gradient_boosting.tiff
│   ├── fig03a_shap_bar.tiff
│   ├── fig03b_shap_beeswarm.tiff
│   ├── fig04a_loho_r2.tiff
│   ├── fig04b_loho_vs_random.tiff
│   ├── fig05a_dist_by_host.tiff
│   ├── fig05b_dist_by_synthesis.tiff
│   ├── fig06a_residuals_vs_predicted.tiff
│   ├── fig06b_residual_histogram.tiff
│   ├── fig06c_cumulative_error.tiff
│   ├── fig07a_conformal_band.tiff
│   ├── fig07b_conformal_coverage.tiff
│   ├── fig08a_noise_bar_chart.tiff
│   ├── fig08b_noise_summary.tiff
│   ├── fig09a_slope_scatter.tiff
│   ├── fig09b_per_host_correlation.tiff
│   ├── fig10a_conc_trend.tiff
│   ├── fig10b_mismatch_trend.tiff
│   ├── requirements.txt
│   └── bandgap_tables.docx                 # manuscript tables (Tables 1–5, S1)
├── Raw_Bandgap_Dataset.xlsx                # literature-extracted dataset before cleaning
├── Doped_Metal_Oxide_Bandgap_Dataset.xlsx  # final 638-point dataset used for modeling
├── code_bandgap.py                         # preprocessing, model training, SHAP, LOHO analysis
├── bandgap_outputs.zip                     # compressed archive of all outputs
└── README.md

Summary of key results

The best-performing model (Gradient Boosting) achieved an MAE of 0.1699 eV on the full dataset, which is almost identical to the median inter-laboratory measurement variation (0.178 eV) computed independently from repeated measurements in the literature. After removing the most inconsistent data pairs, R² improved to 0.83 without any change to the model itself. Leave-One-Host-Out validation shows the model generalizes poorly across different oxide hosts and performs better as an interpolator within a host it has already seen. A within-paper analysis, where data were grouped by source to remove inter-lab noise, showed the model captures a real (if modest) concentration-dependent trend (r = 0.234, p = 0.048). SHAP analysis indicates that synthesis-related parameters are the most influential features overall.

The overall takeaway is that, for this problem, the model's accuracy is already close to the ceiling set by how consistently the underlying experiments were reported — so improving synthesis reporting standards in the literature is likely to matter more for future progress than further algorithmic tuning.

Reproducing the results


Clone this repository.
Install dependencies:


bash   pip install -r bandgap_outputs/requirements.txt


Place Raw_Bandgap_Dataset.xlsx and Doped_Metal_Oxide_Bandgap_Dataset.xlsx in the same directory as code_bandgap.py.
Run the analysis script:


bash   python code_bandgap.py

All figures and tables will be written to bandgap_outputs/.

Citation

If you use this dataset or code, please cite:


Naimur Rahman, Md. Shakil Ahamed, Naim Ferdous, Md. Rahat Al Hassan, Harunur Rashid, and Md Shahabul Alam, "Inter-Laboratory Noise as the Dominant Bottleneck in Machine Learning Prediction of Dopant-Induced Bandgap Shifts in Metal Oxides," [Journal], 2026. [DOI / link once available]



License

This repository is licensed under the MIT License. Copyright (c) 2026 Naimur Rahman.

Contact

Questions or issues with the dataset or code can be raised via GitHub Issues, or contact the corresponding authors:


Naim Ferdous — naim.ferdous.badhan@gstu.edu.bd
Md. Rahat Al Hassan — rahatmme09@cme.ruet.ac.bd
Naimur Rahman — naimur.me07@gmail.com
