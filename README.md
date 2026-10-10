# VS.NET Syntax Word Add-in

Native C# VSTO add-in for desktop Microsoft Word on Windows, targeting .NET Framework 4.8. No Office.js, browser, or web server is required.

## Build and Run

Install Visual Studio with the Office/SharePoint development workload and desktop Word.

1. Run `pwsh -NoProfile -File tools/Build.ps1` to generate a local development signing certificate, restore packages, and build signed VSTO manifests.
2. Open `VsNetSyntaxWordAddIn.sln` in Visual Studio.
3. Press F5 to register the development add-in and launch Word.
4. Select plain code text. In the **VS.NET Syntax** tab, choose **C#**, **HTML**, **TypeScript**, or **JavaScript** from **Syntax Highlight**. Changing the language applies formatting immediately.
5. Use **Apply** to format another selection with the current language.

Formatting modifies character properties in place, preserving text, tabs, line breaks, and the selection. Changes are grouped into one Undo action. The formatter uses Consolas 10 pt on white and rejects protected documents and selections containing cell markers, fields, inline objects, or content controls.

## Highlighting

ColorCode.Core 2.0.15 supplies all four language grammars and its DefaultLight palette. This is Visual Studio-style lexical highlighting, not an exact reproduction of Visual Studio 2026 semantic classification or its current default theme. Word selections have no source-project context, and ColorCode does not use the Visual Studio language services. Exact VS 2026 parity remains outstanding.

## Validation

The native add-in and signed VSTO manifests have built successfully. The test sources in `tests` check token coverage and provide an optional Word integration test:

```powershell
msbuild tests/SmokeTests.csproj
tests/bin/SmokeTests.exe
tests/bin/SmokeTests.exe --word
```

The optional integration test creates a new temporary Word instance and verifies text preservation, colors, selection boundaries, and grouped Undo. Tests have not yet been run successfully.

## Deployment

The build produces `bin/Debug/VsNetSyntaxWordAddIn.vsto`. For distribution, use Visual Studio Publish with your own trusted code-signing certificate. The generated development certificate is for local testing only. Do not commit `DevelopmentSigning.pfx` or `DevelopmentSigning.props`.
