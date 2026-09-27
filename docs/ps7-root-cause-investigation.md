# PowerShell 7 OneNote COM root-cause investigation

The managed proxy bypasses a failing call path; the root cause remains
unresolved. These local probes used Windows, PowerShell 7.6.5 / .NET 10.0.11,
Windows PowerShell 5.1, and Microsoft's OneNote interop assembly 15.0.0.0.
Results are not a diagnosis for every installation. See [Microsoft DLL
architecture] for the types involved.

## Observed failures

- `New-Object -ComObject OneNote.Application` succeeded, but direct PS 7 method
  calls failed with `0x80004005`. The stack included
  `ComRuntimeHelpers.GetITypeInfoFromIDispatch` and
  `IDispatchComObject.EnsureScanDefinedMethods` in
  `System.Management.Automation`. `Get-Member` found neither `GetHierarchy` nor
  `GetSpecialLocation` on the raw object in PS 7; both were discoverable in PS
  5.1.
- A custom `InterfaceIsIDispatch` declaration failed with
  `0x8002801D (TYPE_E_LIBNOTREGISTERED)` in **both** runtimes.
- DLL-free `Type.InvokeMember` calls on the raw object failed with
  `TYPE_E_LIBNOTREGISTERED` in PS 7, using either `GetSpecialLocation` or
  `[DispID=1610743826]`, with explicit by-reference parameter modifiers.
  `$application.GetType()` itself triggered the failing COM binder; the probe
  obtained the type through `[object].GetMethod('GetType').Invoke(...)`.
- Microsoft's `ApplicationClass` could be constructed and its methods listed,
  but direct PowerShell invocation of `GetSpecialLocation` failed. Whether
  execution entered the managed wrapper body was not established.

Occasional activation failures with `0x800706BA` were separate RPC failures, not
evidence about the type-information failure. Word COM member discovery and
reading `Version` worked in the same PS 7 environment.

## Casting and typed calls

PowerShell typed-variable and inline casts to our COM interface failed during
conversion. On the same COM instance, a C# cast and typed call succeeded;
returning the C#-cast reference to PowerShell did not fix normal method calls.
PowerShell casting also failed with an interface created using
`Reflection.Emit`.

Correctly ordered `InterfaceIsDual` declarations worked through compiled C# in
both runtimes. The PS 7 probe needed no Microsoft interop DLL. PS 5.1 loaded an
installed copy automatically during activation, so this was not a DLL-free PS
5.1 result. A successful COM cast checks interface identity, not the correctness
of every declared method slot or marshaling signature.

## Explicit reflection

`MethodInfo.Invoke` from Microsoft's native `IApplicationCOM` worked on the raw
COM object. Methods selected from the managed `IApplication` instead failed with
a target-type mismatch; these interfaces are not interchangeable.

DLL-backed reflection and the C# proxy both passed temporary-notebook creation,
writing, reading, closing, and filesystem cleanup in PS 5.1 and PS 7. Reflection
also worked without the DLL or C# compilation using an `InterfaceIsDual` type
built with `Reflection.Emit`. That fresh-process PS 7 probe covered only
`GetSpecialLocation`, using the first 19 method slots; it did not exercise the
notebook lifecycle or the complete API.

These paths supply known interface metadata instead of asking PowerShell to
discover the OneNote method dynamically. .NET still performs COM marshaling.

## Remaining questions

1. Compare activation identity, process architecture, apartment, loaded
   assemblies, and registered type-library references between matched runs.
2. Inspect `IDispatch.GetTypeInfo` results and type-library registration
   read-only, and trace the failing PowerShell binder path against PS 5.1.
3. Determine whether `ApplicationClass` fails before or inside its managed
   method body; do not assume it shares the raw object's failure mechanism.

The evidence does not yet identify a PowerShell, .NET, or OneNote defect,
establish a registration repair, or rule out every dispatch-based approach.

[Microsoft DLL architecture]: architecture.md
