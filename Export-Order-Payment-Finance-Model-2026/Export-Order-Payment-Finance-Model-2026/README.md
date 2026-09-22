# International Business — Structuring the Payment and Finance of an Export Order

**Internship Studio · International Business Internship · Module 6 Project** — Eswar Mahalingam

| File | What it is |
|---|---|
| `report/Export_Payment_Finance_Report_Eswar.pdf` | 6-page written report: Tasks A–F, executive recommendation, model guide (navy/gold) |
| `model/Export_Payment_Finance_Model_Eswar.xlsx` | Excel model — 7 sheets, 249 live formulas, 0 errors, every cell referencing the Assumptions block |
| `model/Export_Payment_Finance_Model_Eswar_print.pdf` | Landscape print of the workbook (7 pp) for PDF-only portals |
| `model/build_model.py` | openpyxl script that generated the workbook (reproducible) |
| `index.html` | Live dashboard (GitHub Pages: https://eswar5313.github.io/Export-Order-Payment-Finance-Model-2026/) — results tiles, ranking, UCP check, links |
| `brief/` | The original project brief |

## Results (USD)
| Term | Net realisation | PV @12% | All-in cost | Default p | Rank (PV) |
|---|---|---|---|---|---|
| Advance (100%, −5% price) | 569,900 | 569,900 | 5.02% | 0.0% | 2 |
| **Confirmed LC at sight** | **591,330** | **579,735** | **3.38%** | 0.2% | **1 — recommended** |
| Usance LC 90 d (unconfirmed) | 578,580 | 567,235 | 5.46% | 1.0% | 3 |
| D/A 90 d | 561,030 | 550,029 | 8.33% | 4.0% | 4 |
| Open account 90 d | 537,780 | 527,235 | 12.13% | 8.0% | 5 |

Packing credit (Task A): USD 420,000 × 6.0% × 60/360 = **USD 4,200**. Advance beats the sight LC only if the concession is below 3.36%. Task E: 4 UCP 600 discrepancies (invoice 605,000; received-for-shipment B/L; 100% insurance; late shipment) → sight LC falls from rank 1 to rank 4 (all-in cost 3.38% → 11.33%).

Assumptions I added (yellow cells): 2-month confirmation tenor; discrepant-LC default probability 8%; 10-day waiver delay.

## Submit
Folder `Eswar_Mahalingam_IB_Module6_Project` = report PDF + xlsx → Google Drive → "Anyone with the link can view" → paste link in the Internship Studio submission portal.
