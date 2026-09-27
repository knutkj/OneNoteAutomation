# OneNote interop generator

The [proxy generator] maintains the C# interop source used by the module's
PowerShell 7 wrapper. It reads Microsoft's OneNote interop assembly and emits
COM declarations and forwarding methods, including the convenience overloads and
their default arguments.

Generation is a development operation. The module uses the checked-in C# source,
compiled with `Add-Type` on first use in PowerShell 7; users do not need the
generator or Microsoft's interop DLL. Windows PowerShell 5.1 continues to use
the raw COM object.

## Input and output

- **Input:** a local copy of `Microsoft.Office.Interop.OneNote.dll`. The
  generator was developed against assembly version `15.0.0.0`.
- **Output:** [OneNoteInterop.Generated.cs], in the `OneNoteAutomation.Interop`
  namespace. Each run overwrites this file.
- **Handwritten code:** [OneNoteApplication.cs] handles COM activation and
  disposal. The generator does not modify it.

The generator reads `IApplicationCOM` for the native contract and follows its
referenced interfaces, enums and structures. It reads `IApplication` and the
forwarding bodies in `ApplicationClass` for the additional overloads and their
defaults. See [Microsoft DLL architecture] for the source assembly's type
inventory.

## Regenerate

From a PowerShell session at the repository root, supply the path to your
development copy of the DLL:

```powershell
./scripts/New-InteropProxy.ps1 `
    -InteropAssemblyPath C:/path/to/Microsoft.Office.Interop.OneNote.dll
```

Generation reads assembly metadata and method bodies; it does not activate
OneNote or access notebooks. The script reports the source assembly identity,
output path and generated member counts.

Review the output before committing it:

```powershell
git diff -- OneNoteAutomation/Interop/OneNoteInterop.Generated.cs
Invoke-Pester -Path ./tests
```

Run the tests in both Windows PowerShell 5.1 and PowerShell 7 on Windows. They
do not replace review of native method order and marshaling declarations. Do not
edit the generated file manually; change the generator and regenerate.

## Boundaries

This is a OneNote-specific generator, not a general-purpose COM importer.
Unsupported types, marshaling forms or overload forwarding bodies cause it to
fail rather than guess. Changes to the input DLL may require generator changes.

Matching method signatures does not reproduce all behavior of Microsoft's
wrapper. The generator does not implement its automatic busy retries or event
support, and returned COM subobjects do not receive their own managed wrappers.

[proxy generator]: ../scripts/New-InteropProxy.ps1
[OneNoteInterop.Generated.cs]:
  ../OneNoteAutomation/Interop/OneNoteInterop.Generated.cs
[OneNoteApplication.cs]: ../OneNoteAutomation/Interop/OneNoteApplication.cs
[Microsoft DLL architecture]: architecture.md
