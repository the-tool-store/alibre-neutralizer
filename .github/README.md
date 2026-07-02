# Alibre Neutralizer

Alibre Neutralizer is a configurable, repeatable bulk exporter for Alibre Design.

It walks a top-level assembly and exports every subassembly and part to neutral CAD exchange formats, descending through nested assemblies. Each run reads its settings from an XML configuration file, so repeated exports produce identical output. That output fits Git-versioned hardware projects. Each component is exported once.

The tool runs two ways: as an Alibre Script inside the Alibre Script add-on, or as a compiled C# add-on that registers its own ribbon menu and runs the bundled script through IronPython. The add-on project references Alibre Design 29.0.0.29060 assemblies, targets .NET Framework 4.8.1 and x64, and uses IronPython 2.7.12.

## Table Of Contents

- [What Is Here](#what-is-here)
- [Official Alibre Resources](#official-alibre-resources)
- [Requirements](#requirements)
- [Quick Start](#quick-start)
- [Installation](#installation)
- [Usage](#usage)
- [Key Files](#key-files)
- [Key Folders](#key-folders)
- [Screenshots](#screenshots)
- [Notes](#notes)
- [License](#license)

## What Is Here

- Recursive export of a complete assembly: every subassembly and part becomes an individual file, and each component exports only once.
- Neutral formats: STEP (AP203 and AP214), SAT, IGES, and STL.
- Metadata sidecars in CSV: Alibre Properties (`CSV_Properties`) and Design Parameters / equations (`CSV_Parameters`). Parameters are alphabetized so output stays stable across runs and produces consistent version-control diffs.
- Parametric file and folder naming from Alibre Properties (for example `{Number}`, `{Name}`, `{Supplier}`, `{Revision}`).
- Multiple Export Directives in a single pass, each with its own format, path scheme, and rules for whether the root assembly, subassemblies, and parts are included.
- Optional pre-export purge that clears only the matching file types from a target directory before writing fresh exports.
- A standalone IronPython script plus an optional C# add-on with an Inno Setup installer.

## Official Alibre Resources

Alibre's official resources for API development and AI/LLM/agent workflows: <https://www.alibre.com/api/>

## Requirements

- Alibre Design. The C# add-on project references Alibre Design 29.0.0.29060 assemblies and targets x64.
- To run the script directly: Alibre Design with the Alibre Script add-on.
- To build the optional add-on: .NET Framework 4.8.1, IronPython 2.7.12, and Inno Setup for the installer.

## Quick Start

Run it as an Alibre Script. Copy `source/alibre-neutralizer.py` into a subfolder of your Alibre Script Library, open the top-level assembly in Alibre Design, run the script from the Alibre Script add-on, and select an XML configuration file when prompted. Use `source/example-alibre-neutralizer-config.xml` as a starting template.

## Installation

Alibre Neutralizer works two ways.

**As an Alibre Script:**
Download or clone the repository into a subfolder of the Alibre Script Library (by default under the Documents folder, for example `C:\Users\<user>\Documents\Alibre Script Library`). Only `source/alibre-neutralizer.py` is required; keeping the full repository also provides the example configuration file.

**As a compiled add-on:**
The `source/alibre-neutralizer-addon/` folder contains a C# add-on (`.adc` manifest plus IronPython host) that registers an "Alibre Neutralizer" ribbon menu and runs the bundled script. Build `source/alibre-neutralizer-addon/alibre-neutralizer-addon.sln`, then run the Inno Setup script (`alibre-neutralizer-addon.iss`) to produce an installer that copies the add-on into the Alibre `Addons` directory and registers it under the `Alibre Design Add-Ons` registry key.

## Usage

1. Create an XML configuration file describing your Export Directives. Each directive sets a `type` (`STEP203`, `STEP214`, `SAT`, `STL`, `IGES`, `CSV_Properties`, or `CSV_Parameters`), a `RelativeExportPath` (with optional `{Property}` placeholders), an optional `PurgeDirectoryBeforeExporting` path, and the `EnableRootAssemblyExport` / `EnableSubassemblyExport` / `EnablePartExport` flags. See `source/example-alibre-neutralizer-config.xml` for a working template.
2. Open the top-level assembly you want to export in Alibre Design.
3. Run the tool: in the Alibre Script add-on, open `source/alibre-neutralizer.py` and click Run; or, if the add-on is installed, select "Run Alibre Neutralizer" from the Alibre Neutralizer ribbon menu.
4. Select your configuration file when prompted, review the summary dialog (which reports how many export directives were parsed), and confirm to start the export.
5. Progress and any errors are logged to the Alibre Script console; a notification appears when the export completes.

Paths resolve relative to the configuration file's location, offset by the optional `BaseExportPath`. Exporting from Alibre PDM is unreliable; export a package and run against that instead.

## Key Files

| File | Purpose |
| --- | --- |
| `source/alibre-neutralizer.py` | Main export script; the only file required to run as an Alibre Script. |
| `source/example-alibre-neutralizer-config.xml` | Example configuration with Export Directives to copy and adapt. |
| `source/AlibreScript.py` | Alibre Script API stub for editor autocomplete and type checking during development. |
| `source/alibre-neutralizer-addon/alibre-neutralizer-addon.sln` | Visual Studio solution for the C# add-on. |
| `source/alibre-neutralizer-addon/alibre-neutralizer-addon.iss` | Inno Setup script that builds the add-on installer. |
| `source/alibre-neutralizer-addon/alibre.disclaimer.txt` | Disclaimer text bundled with the add-on. |
| `source/alibre-neutralizer-addon/src/AlibreAddOn.cs` | C# add-on host that registers the ribbon menu and runs the script. |
| `source/alibre-neutralizer-addon/src/alibre-neutralizer-addon.adc` | Add-on manifest. |
| `source/alibre-neutralizer-addon/src/alibre-neutralizer-addon.csproj` | C# project file for the add-on. |
| `source/alibre-neutralizer-addon/src/Scripts/alibre-neutralizer.py` | Copy of the export script bundled with and run by the add-on. |

## Key Folders

| Folder | Purpose |
| --- | --- |
| `source/` | Export script, example configuration, API stub, and the add-on project. |
| `source/alibre-neutralizer-addon/` | C# add-on: solution, Inno Setup installer script, and source. |
| `source/alibre-neutralizer-addon/src/` | Add-on C# host, manifest, project file, and bundled script. |
| `source/alibre-neutralizer-addon/src/Scripts/` | Bundled export script the installed add-on runs. |
| `documentation/` | Screenshots referenced in this README. |
| `reviews/` | Code review notes. |
| `submodules/` | Placeholder for submodules; currently empty. |
| `.github/` | This README. |

## Screenshots

Run the script from the Alibre Script ribbon:

![Run the script](../documentation/step-1-run-script.png)

Select your configuration file:

![Select configuration](../documentation/step-2-select-configuration.png)

Confirm the export:

![Confirm export](../documentation/step-3-confirm-export.png)

Progress is logged to the Alibre Script console:

![Console output](../documentation/step-4-console.png)

Completion notification:

![Successful export notification](../documentation/successful-export-notification.png)

## Notes

- Alibre Script runs on IronPython 2.7, so the script stays Python 2.7 compatible.
- Design Parameters export alphabetized, which keeps CSV diffs stable across runs.
- Each component exports only once per pass, even when it appears in multiple subassemblies.
- Exporting from Alibre PDM is unreliable; export a package and run against that.

## License

GNU Lesser General Public License v3.0. See [LICENSE](../LICENSE.md).
