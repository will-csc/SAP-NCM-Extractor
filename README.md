# SAP NCM Extractor

### Excel VBA + SAP GUI Scripting

An Excel-based automation tool developed using **VBA and SAP GUI Scripting** to retrieve NCM (Mercosur Common Nomenclature) classification codes from SAP Material Master Data across multiple plants.

The project eliminates repetitive manual consultations through SAP transaction `MM03`, centralizing the extracted information in a structured Excel workbook.

## Features

- Automated integration between Excel and SAP GUI.
- Batch processing of multiple material codes.
- NCM extraction through transaction `MM03`.
- Retrieval across three configurable plants.
- Automatic organization of results in Excel.
- Reduced manual navigation and repetitive tasks.

## Technologies

| Technology | Purpose |
|---|---|
| Excel VBA | Automation logic |
| SAP GUI Scripting | SAP interaction |
| SAP MM | Material Master Data |
| MM03 | Material information retrieval |

## How It Works

The automation follows five main steps:

1. Connects to an active SAP GUI session.
2. Reads material codes from column A of the Excel worksheet.
3. Navigates through transaction `MM03`.
4. Retrieves the classification from SAP field `MARC-STEUC` for each configured plant.
5. Automatically populates the corresponding Excel columns.

### Example Output

| Material Code | Plant 8801 | Plant 8802 | Plant 8800 |
|---|---|---|---|
| MAT-001 | 85044090 | 85044090 | 85044090 |
| MAT-002 | 85371090 | 85371090 | 85371090 |
| MAT-003 | 73089010 | 73089010 | 73089010 |

*Illustrative data only.*

## Requirements

- Microsoft Excel Desktop with VBA support.
- SAP GUI for Windows.
- SAP GUI Scripting enabled.
- Active SAP session with authorization for `MM03`.
- Access to the relevant material and plant data.

## Installation & Usage

**1.** Open Excel and create a macro-enabled workbook (`.xlsm`).

**2.** Insert material codes into column A, starting from `A2`.

**3.** Open the VBA Editor (`ALT + F11`), insert a module, and paste the source code.

**4.** Configure your SAP plant identifiers:

```vb
For Each plant In Array("8802", "8801", "8800")
```

Update the output-column mapping if necessary.

**5.** Log in to SAP GUI and execute the `getNCM` macro.

The extracted NCM classifications will automatically populate the worksheet.

## Technical Details

The connection to SAP is established through SAP GUI Scripting:

```vb
Set SapGuiAuto = GetObject("SAPGUI")
Set Application1 = SapGuiAuto.GetScriptingEngine
Set Connection = Application1.Children(0)
Set session = Connection.Children(0)
```

The automation retrieves the classification from `MARC-STEUC`, whose use for NCM depends on the SAP localization and configuration.

## Limitations & Future Improvements

The current implementation depends on SAP GUI screen layouts and requires an active SAP session.

Potential enhancements include:

- Dynamic plant selection.
- Error logging and input validation.
- Execution progress tracking.
- Automatic SAP session identification.
- A user-friendly Excel interface.

## Disclaimer

This project demonstrates process automation using Excel VBA and SAP GUI Scripting.

It requires valid SAP credentials and appropriate authorizations. Users should comply with their organization's security policies and avoid exposing confidential corporate information.
