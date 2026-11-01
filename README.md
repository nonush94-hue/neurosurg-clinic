# Neurosurg Exam

Bedside tool for neurosurgical visits (Italian UI): history, presentation, therapy, condition-specific data and the neurological exam. It generates a note for the clinical record and keeps a research database you can export to Excel.

- Patient archive on the device (IndexedDB): pre-op, post-op, discharge and follow-up visits per patient
- 24 condition categories (pituitary adenoma, glioma, meningioma, metastases, aneurysm, AVM, chronic SDH, hydrocephalus, disc herniation, myelopathy, spondylolisthesis…) with their own scales, defined in `fields.js`
- Excel export: Anagrafica sheet, Visite sheet, one sheet per condition. Research sheets identify patients by code only. An anonymous export leaves out the Anagrafica sheet.
- JSON backup and restore
- Exam guide and glossary; works offline and installs on iPad (Share → Add to Home Screen)
- No server and no tracking: patient data never leaves the device unless you export it

Files: `index.html` (layout and styles), `app.js` (logic), `fields.js` (sections and conditions), `xlsx.js` (Excel writer), `sw.js` (offline cache; bump the cache name when CORE files change).

Part of the Neurosurg family: Neurosurg Predict, Neurosurg Learn, Neurosurg Exam.
