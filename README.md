# E-XGBoost: Early Undergraduate Dropout Prediction in Ghanaian Higher Education

An engineered gradient-boosting model for **early prediction of undergraduate dropout** at a Ghanaian
university, using Moodle LMS engagement logs and survey data.

**E-XGBoost** applies two model-engineering operations to standard XGBoost:

- **[F] Fabricate** — a class-balanced **focal-loss objective** (custom gradient + Hessian) replacing the
  default log-loss, to target severe class imbalance.
- **[R] Remove** — **SHAP-ranked feature pruning** (drop the bottom ~30% by mean |SHAP|), reducing model
  size while preserving predictive power.

This repository accompanies the Q1-journal research proposal (Target: *Computers & Education*, Elsevier).

## Research questions

- **RQ1** — Can a fabricated focal-loss objective with SHAP-guided pruning be integrated into XGBoost?
- **RQ2** — Does E-XGBoost significantly improve AUC-PR and recall over a locked XGBoost baseline?
- **RQ3** — Which LMS/survey features most drive dropout risk, and does pruning preserve performance?

## Method (reproducibility)

- Data: hybrid — primary KNUST Moodle logs + survey (CHRPE ethics clearance) + public benchmark.
- Split: stratified 80/20, then 5-fold stratified CV. Fixed `SEED = 42`.
- Baseline (locked): standard XGBoost (log-loss). Engineered: E-XGBoost on identical splits.
- Significance: Wilcoxon Signed-Rank Test on fold-level AUC-PR; effect size = Δ AUC-PR.

## Repository layout

```
generate_proposal.py   # builds the proposal .docx
requirements.txt       # pinned dependencies
src/                   # (planned) baseline + E-XGBoost notebooks/scripts
data/                  # (gitignored) local data — never committed
```

## Status

Proposal stage. Ethics application (KNUST CHRPE) in progress; data collection pending approval.

## Repository

https://github.com/Notorious1x/e-xgboost-dropout

## Author

Adams Caleb Yeboah Yaw — Department of Computer Science, KNUST — pokaa96@gmail.com

## License

MIT
