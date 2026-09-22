# Internship-Studio-Projects-2026

**Eswar Mahalingam** · Multi-domain internship, Internship Studio (batches 11 Apr · 15 Apr · 22 Apr · Aug 2026) · Ghaziabad, India
📊 **Live dashboard:** https://eswar5313.github.io/Internship-Studio-Projects-2026/ · 📧 eswarmba05313@gmail.com · 📱 +91-9360548243

One numbered folder per subject. Every folder is standalone-ready and follows the same layout: `brief/` (original project brief) → deliverables (workbook / notebook / code / design files) → a navy/gold first-person report PDF → `README.md`. All workbooks recalculate with **0 formula errors**; every declared-synthetic dataset is labelled as such.

## Projects (18)

| # | Subject | Project | Key result |
|---|---|---|---|
| [01](01_AI/) | Artificial Intelligence | Project Panopticon — exam-proctoring anomaly detection notebook (merge_asof, 10 s rolling windows, balanced RandomForest, 0.90 precision threshold + PR curve) | Precision 1.00 / recall 0.48 at 0.90 threshold — GO |
| [02](02_Data_Science/) | Data Science | Project Overheat — predictive maintenance (12 h rolling features, 24 h target shift, chronological split, balanced RF) | 3/3 test failures warned 23–24 h ahead; top feature Sensor_Vibration_mm 0.388 |
| [03](03_Software_Testing/) | Software Testing | OrderFlow regression prioritisation — APFD, greedy ordering, set-cover, 20-min budget (Excel 3,108 formulas + report) | APFD 75.83% → 89.17%; reduced suite T7,T1,T6,T8 = 16 min (−62.8%) |
| [04](04_CPP_DSA/) | C++ & Data Structures | Smart Library Management System — linked list, stack undo, queue reservations, sort, books.txt loader | Compiles `-Wall -Wextra` clean; valgrind 0 leaks |
| [05](05_Energy_Conservation/) | Energy Conservation & Management | Textile mill HT electricity-bill reconstruction — TOD / PF / MD analysis, 4+ savings measures (1,062 formulas) | Annual bill Rs 10.76 Cr @ Rs 10.26/kWh; savings Rs 81.9 L (7.6%) |
| [06](06_RCC_Design/) | RCC Structure Design | One-way slab 10×4 m, doubly-reinforced hall slab + beam, biaxial column (IS 456, M25/Fe415 & M20/Fe415) | Slab D170 10@120; beam 400×750 (8-32 + 8-25); column 300×450 8-16, interaction 0.389 |
| [07](07_CartShare/) | Web (CartShare) | Collaborative cart web app — HTML/CSS/JS, GitHub + live deploy | 54/54 Playwright checks passing |
| [08](08_AutoCAD/) | AutoCAD | Civil single-storey plan, electrical amplitude-limiter schematic + wiring schedule, mechanical screw-jack 7-part sheet + BOM (DXF via ezdxf) | 1,230 entities, audit 0 errors |
| [09](09_Finance/) | Finance | Ravi D/E case, Vijay 10% share issue, Aurora Consumer Foods 2-year ratio analysis | D/E 1.50 → 1.92; dilution 9.09%; verdict BUY/ACCUMULATE |
| [10](10_Excel_Automation/) | Excel Automation | Multi-region sales dashboard on a declared-synthetic 645-row dataset | 12,621 formulas; Rs 14.26 Cr revenue = 101% of target; South only region under target |
| [11](11_HR/) | Human Resources | ABC Co. 18-page policy manual — 3-day induction, 90-day onboarding, discipline ladder, POSH/maternity, engagement proposal + workbook | Engagement budget Rs 4,00,000 = Rs 5,714/employee (0.95% of payroll) |
| [12](12_Machine_Learning/) | Machine Learning | Student Performance model — Ridge Regression on the real StudentsPerformance dataset (notebook + script) | R² 0.8805, MAE 4.21 |
| [13](13_AWS/) | AWS | Express gallery app on EC2 + S3, ROS Noetic deployment runbook, 11 CLI/bootstrap scripts | 7/7 local tests; t2.micro needs 2 GB swap |
| [14](14_Corporate_Psychology/) | Corporate Psychology | "TechX Thrive" well-being proposal — 5 pillars, 27 interventions, 20 KPIs, pilot-vs-comparison DiD evaluation | Year-1 budget Rs 1.21 Cr = Rs 23,341/employee |
| [15](15_Graphic_Design/) | Graphic Design (Illustrator) | 34 original SVGs — FinPay icon set, "Ride the Grind" tee, Save Water posters, patterns, LearnLoop mascots, Brew Street Café social kit, HUSTLE 3D lettering | 10-page portfolio PDF |
| [16](16_Digital_Marketing/) | Digital Marketing | EcoStyle Fashion targeted Facebook ad campaign — workbook + 5 creatives + 12-page report | ROAS 4.53×, CPA Rs 331, +23.5% sales |
| [17](17_International_Business/) | International Business | Export order payment & finance model — 5 payment terms, PV @ 12% cost of capital, UCP 600 discrepancy check, EXIM Bank / WTO / World Bank analysis (Excel 249 formulas + 6-page report) | **Recommend confirmed sight LC** — PV USD 579,735, all-in cost 3.38%; advance beats it only if concession < 3.36% |
| [18](18_Video_Editing/) | Video Editing | Reel production blueprint — Idea 4 shop intro ("Kulhad & Co."), 48 s @ 96 BPM timeline, Premiere Pro stage map (01–07 + 5 bonus), cover-frame template, sources credit log | 16/16 techniques + 5/5 bonus planned; reel shot & edited by me in Premiere Pro |

## Repo layout
```
Internship-Studio-Projects-2026/
├── index.html            ← dashboard (GitHub Pages, root)
├── dashboard/index.html  ← same dashboard
├── 01_AI/ … 18_Video_Editing/
│   ├── brief/            original project brief
│   ├── <deliverables>    workbook / notebook / src / model / docs
│   ├── NN_Report.pdf     navy/gold first-person report
│   └── README.md
└── README.md
```

## Standards applied to every folder
- **Reproducible:** build scripts included wherever a file was generated (openpyxl, ezdxf, notebooks).
- **Live formulas, zero errors:** every Excel model recalculated with LibreOffice before commit; inputs blue, editable assumptions yellow.
- **Declared substitutions:** where a brief's dataset or tool was unavailable, the substitute is stated on the report cover (e.g. synthetic sales rows, Python equivalents).
- **No fabrication:** real datasets are cited; synthetic data is labelled; interview/user-data slots stay blank until filled with real inputs.
- **Reports fill the page:** navy/gold sidebar layout, name / course / assignment on page 1.

## Submission
Each project is submitted to Internship Studio as a Google Drive link ("Anyone with the link can view") on the Project Submission Dashboard / iStudio assignments portal. Folder 18's final MP4 + `.prproj` are added to `18_Video_Editing/submission/` after the Premiere Pro export.

## Related repos
- [Export-Order-Payment-Finance-Model-2026](https://github.com/Eswar5313/Export-Order-Payment-Finance-Model-2026) — standalone version of project 17
- [Reel-Editing-Project-2026](https://github.com/Eswar5313/Reel-Editing-Project-2026) — standalone version of project 18
- [Codec-Technologies-Internship-Portfolio-2026](https://github.com/Eswar5313/Codec-Technologies-Internship-Portfolio-2026) · [GraySentinel-DSOU-45Day-2026](https://github.com/Eswar5313/GraySentinel-DSOU-45Day-2026)

---
Eswar Mahalingam · Ghaziabad, India · 2026
