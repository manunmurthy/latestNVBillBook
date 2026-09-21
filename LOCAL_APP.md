# Local Bill Book App

This is the everyday, no-command interface for generating Nava Vaibhva reports. It runs only on your computer at `http://127.0.0.1:5000`; bank statements and reports do not leave the computer.

## First-time setup

1. Keep `config.yaml`, `flats.yaml`, and `flat_dimensions.yaml` in the main project folder. `config.yaml` is already present in the working installation.
2. Open Terminal in the project folder and run the following once:

   ```bash
   source .venv/bin/activate
   pip install -r requirements.txt
   chmod +x start_billbook.command
   ```

## Starting and stopping the app

Double-click `start_billbook.command`. A Terminal window opens and starts the app. Then open the address shown in that window, normally:

```text
http://127.0.0.1:5000
```

When you have finished, return to that Terminal window and press `Control + C`. Closing that Terminal window also stops the app. Reports already generated remain safely in `data/output/`.

## Monthly bill report

1. In **Monthly bill report**, select one or more downloaded HDFC statement files.
2. Enter the water bill per flat when it differs from the configured amount. Otherwise leave it blank.
3. Select **Generate monthly report**.
4. Choose **View first** to check the workbook in the browser, or **Download Excel** to save/open the complete workbook.

A full monthly report is created for every flat in the registry, exactly as with the previous monthly process. Multiple selected statements result in one monthly workbook per statement.

## Collection history: one flat, many flats, or every flat

1. In **Collection history**, select the statement files covering the required period. You may select one multi-month file or several monthly files together.
2. In **Flat number(s)**, enter a single flat such as `A001`, a comma-separated list such as `A001, A002, B303`, or paste one flat per line.
3. Leave the flat field empty to include every flat.
4. Generate, preview, and download the workbook.

The app validates flat numbers against `flats.yaml`. Before making a history report, ensure `meta.water_bills_by_month` in `flats.yaml` has a water-bill value for every month in the selected period.

## Where reports are saved

Every generated Excel workbook is retained in:

```text
data/output/
```

The browser preview is deliberately limited to the first 100 rows and 25 columns of a selected worksheet, so it stays quick and easy to read. The downloaded Excel workbook always contains the complete report and all worksheets.

## Troubleshooting

- **The app says it cannot find `config.yaml`:** start it from the project folder using `start_billbook.command`, and check that `config.yaml` exists beside that file.
- **A flat is not found:** correct the flat number or add it to `flats.yaml` before generating the history report.
- **A water-bill message appears for history:** add the missing month(s) under `meta.water_bills_by_month` in `flats.yaml`.
- **The page will not open:** check that the Terminal window is still open and that it says the app is ready. Then visit `http://127.0.0.1:5000` exactly.
