# DPHE ML literature review — second-stage verified version (5 Oct 2026)

| File | Contents |
|---|---|
| `DPHE_ML_Literature_Log_v2_Verified.xlsx` | Verified database: Literature Log, Core ML Evidence, Model Comparison, Research Gaps, Statistical Analysis (KPIs, 18 charts, 14 cross-tabs, key findings), Removed Papers, Download/PDF Verification, Benchmarks & Plan, PDF Rename Map, Legend & Method |
| `DPHE_ML_Literature_Review_v2_Verified.docx` | Literature review rewritten from scratch (23 sections + appendices) |
| `rename_pdfs.ps1` / `rename_pdfs.sh` | Rename the 26 PDFs in `ML_DPHE_Papers` to `P##_Author_Year_ShortTitle.pdf` / `REMOVED_...` (rename only, nothing deleted) |
| `build/` | Source data (`papers_data.py`, `removed_data.py`, `content.py`, `paragraphs.py`), builders, figures, and per-paper extraction notes |

Rebuild: `cd build && python3 build_workbook.py && python3 build_report.py` (needs openpyxl, python-docx, matplotlib).

Known limits: journal quartiles not verified (SCImago/OpenAlex blocked by network policy); OneDrive connector read-only, so PDFs were not renamed in place.
