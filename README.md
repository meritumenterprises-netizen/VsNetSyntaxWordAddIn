# VS.NET Syntax Word Add-in

Native C# VSTO add-in for desktop Microsoft Word on Windows, targeting .NET Framework 4.8. No Office.js, browser, or web server is required.

## Build and Run

Install Visual Studio with the Office/SharePoint development workload and desktop Word.

1. Run `pwsh -NoProfile -File tools/Run.ps1` to generate a local development signing certificate, restore packages, build signed VSTO manifests, register the add-in for your Windows account, and launch a new Word instance.
2. For debugging, open `VsNetSyntaxWordAddIn.sln` from its current folder in Visual Studio. After moving the project, reopen this solution so Visual Studio uses the new paths.
3. Press F5 to register the development add-in and launch Word.
4. Select plain code text. In the **VS.NET Syntax** tab, choose **C#**, **HTML**, **TypeScript**, or **JavaScript** from **Syntax Highlight**. Changing the language applies formatting immediately.
5. Use **Apply** to format another selection with the current language.

The **Theme** dropdown offers **Light** (the default) and **Dark** for language text colors. The **Background** dropdown offers **Normal** (the default, leaves existing background formatting untouched), **Red** (`#FED9B4`), and **Green** (`#E2EFD9`), with rectangular color swatches. Office places dropdown item images to the left of their labels; right-side images are not supported by this native control. Changing either dropdown immediately reformats the selection. Apply and keyboard shortcuts use both current settings. Background colors apply only to selected code, leaving the document page color unchanged.

Formatting modifies character properties in place, preserving text, tabs, line breaks, and the selection. Changes are grouped into one Undo action. The formatter uses Consolas 10 pt and rejects protected documents and selections containing cell markers, fields, inline objects, or content controls.

To build without registering or launching Word, run `pwsh -NoProfile -File tools/Build.ps1`. To build and register only, add `-Register`. The build script locates Visual Studio by its installed MSBuild and OfficeTools targets, avoiding installer component identifiers that differ between installations.

## Ribbon Designer

Open `SyntaxRibbon.cs` in Solution Explorer, right-click it, and select **View Designer** (Shift+F7). Use the Office Ribbon Controls toolbox and Properties window to edit the layout, labels, and icons. The ribbon uses Visual Studio's VSTO Ribbon Designer structure; `SyntaxRibbon.Designer.cs` holds the designer-managed layout and `SyntaxRibbon.cs` holds the event handlers. The old `SyntaxRibbon.xml` is retained for reference only and is not loaded.

## Highlighting

The add-in applies highlighting to the current selection with these shortcuts while a Word document window is active:

| Shortcut | Language |
| --- | --- |
| Ctrl+Shift+Alt+C | C# |
| Ctrl+Shift+Alt+J | JavaScript |
| Ctrl+Shift+Alt+T | TypeScript |
| Ctrl+Shift+Alt+H | HTML |

Both left and right modifier keys work. Holding a shortcut applies it once per key press. The ribbon language changes to match the shortcut. The loaded Word add-in owns these native Windows hotkeys: it registers them while a Word document window is active and releases them when focus leaves Word, the add-in unloads, or Word closes. The old Windows-key shortcuts are no longer registered.

ColorCode.Core 2.0.15 supplies all four language grammars and its DefaultLight and DefaultDark palettes. This is Visual Studio-style lexical highlighting, not an exact reproduction of Visual Studio 2026 semantic classification or its current default theme. Word selections have no source-project context, and ColorCode does not use the Visual Studio language services. Exact VS 2026 parity remains outstanding.

## Validation

The native add-in and signed VSTO manifests have built successfully. The test sources in `tests` check token coverage and provide an optional Word integration test:

```powershell
msbuild tests/SmokeTests.csproj
tests/bin/SmokeTests.exe
tests/bin/SmokeTests.exe --word
```

The syntax smoke tests and optional Word integration test passed locally. The integration test creates a new temporary Word instance and verifies text preservation, colors, selection boundaries, and grouped Undo. Word also loaded the registered development add-in successfully (`COMAddIns.Item("VsNetSyntaxWordAddIn").Connect` was true). These checks do not verify Visual Studio's F5 debugger launch.

## Deployment

The build produces `bin/Debug/VsNetSyntaxWordAddIn.vsto`. For distribution, use Visual Studio Publish with your own trusted code-signing certificate. The generated development certificate is for local testing only. Do not commit `DevelopmentSigning.pfx` or `DevelopmentSigning.props`.
