# FormuRay

FormuRay is a Microsoft Excel VSTO add-in designed to improve formula auditing, model review, navigation, formatting, and consistency checks in financial models and other complex Excel workbooks.

The project is intended for analysts and other Excel power users who work with formula-heavy workbooks and need faster ways to inspect, validate, and navigate models.

## Key Features

- **Formula Audit Explorer** — inspect formulas, referenced elements, values, and formula components.
- **Cross-Sheet Inflow / Outflow** — highlight links coming from other worksheets and links going to other worksheets.
- **Hardcode Detector** — identify hardcoded values in calculation areas.
- **Formula Consistency Map** — highlight changes in formula patterns across rows.
- **Named Ranges Manager** — inspect and manage workbook named ranges, including broken and hidden names.
- **Used Range / Content Zone** — identify the meaningful calculation area of a worksheet.
- **AutoFormatter** — apply consistent formatting to inputs, formulas, percentages, dates, values, borders, and alignment.
- **Paste Formulas** — paste formulas without applying Excel's normal relative-reference shift.

## Platform

FormuRay is built as a Microsoft Excel VSTO Add-in.

### Requirements

- Microsoft Windows
- Microsoft Excel desktop
- .NET Framework 4.6.2 or later compatible runtime
- Microsoft Visual Studio Tools for Office Runtime

Actual deployment requirements may depend on the environment and the installation package used by the organization.

## Deployment

FormuRay can be deployed through a standard Windows installation package or through an organization's software distribution infrastructure.

In a corporate environment, centralized deployment by IT is recommended after the required internal security and software review.

End users may receive FormuRay through the organization's standard software deployment process without direct access to the installation package.

## License

Copyright 2026 Artem Sunguryan

FormuRay is licensed under the **Apache License, Version 2.0**.

You may use, reproduce, modify, and distribute FormuRay in accordance with the terms of the Apache License 2.0.

See the [`LICENSE`](LICENSE) file for the full license text.

The license applies to the version of FormuRay distributed with this package. Future versions of FormuRay may be released under different licensing terms.

## Disclaimer

FormuRay is provided on an **"AS IS" BASIS**, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied, as described in the Apache License 2.0.

Users and organizations are responsible for validating the software and its output before relying on it in production, financial, reporting, or decision-making processes.
