#
# .SYNOPSIS
# Generates the module COM declarations and forwarding methods from OneNote metadata.
#
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$InteropAssemblyPath
)

$ErrorActionPreference = 'Stop'
$assemblyPath = (Resolve-Path -LiteralPath $InteropAssemblyPath).Path
$assembly = [Reflection.Assembly]::LoadFrom($assemblyPath)
$applicationType = $assembly.GetType('Microsoft.Office.Interop.OneNote.IApplicationCOM', $true)
$managedInterface = $assembly.GetType('Microsoft.Office.Interop.OneNote.IApplication', $true)
$managedClass = $assembly.GetType('Microsoft.Office.Interop.OneNote.ApplicationClass', $true)
$types = [Collections.Generic.Dictionary[string, type]]::new()
$pending = [Collections.Generic.Queue[type]]::new()
$pending.Enqueue($applicationType)

function Get-TypeName([type]$Type) {
    if ($Type.IsByRef) { return Get-TypeName $Type.GetElementType() }
    if ($Type -eq [void]) { return 'void' }
    $aliases = @{
        'System.Byte' = 'byte'; 'System.SByte' = 'sbyte'
        'System.Int16' = 'short'; 'System.UInt16' = 'ushort'
        'System.Int32' = 'int'; 'System.UInt32' = 'uint'
        'System.Int64' = 'long'; 'System.UInt64' = 'ulong'
        'System.String' = 'string'; 'System.Boolean' = 'bool'
        'System.Object' = 'object'
    }
    if ($aliases.ContainsKey($Type.FullName)) { return $aliases[$Type.FullName] }
    if ($Type.Assembly -eq $assembly) {
        if ($Type.IsInterface -and $Type.GUID -eq $applicationType.GUID) {
            return 'IOneNoteApplication'
        }
        if (-not ($Type.IsValueType -or ($Type.IsInterface -and $Type.IsImport))) {
            throw "Unsupported metadata type: $Type"
        }
        return $Type.Name
    }
    return $Type.FullName
}

while ($pending.Count) {
    $current = $pending.Dequeue()
    if ($current.IsByRef) { $current = $current.GetElementType() }
    if ($current.Assembly -ne $assembly) { continue }
    if ($current.IsInterface -and $current.GUID -eq $applicationType.GUID) {
        $current = $applicationType
    }
    if ($types.ContainsKey($current.FullName)) { continue }
    $null = Get-TypeName $current
    $types.Add($current.FullName, $current)
    if ($current.IsEnum) { continue }
    if ($current.IsValueType) {
        foreach ($field in $current.GetFields()) { $pending.Enqueue($field.FieldType) }
        continue
    }
    foreach ($method in $current.GetMethods()) {
        $pending.Enqueue($method.ReturnType)
        foreach ($parameter in $method.GetParameters()) {
            $pending.Enqueue($parameter.ParameterType)
        }
    }
}

function Get-MarshalAttribute($Parameter, [switch]$ReturnValue) {
    $marshal = $Parameter.GetCustomAttributes([Runtime.InteropServices.MarshalAsAttribute], $false) |
    Select-Object -First 1
    if ($null -eq $marshal) { return '' }
    $target = if ($ReturnValue) { 'return: ' } else { '' }
    if ($marshal.Value -eq [Runtime.InteropServices.UnmanagedType]::CustomMarshaler) {
        $marshalType = $marshal.MarshalType.Replace('\', '\\').Replace('"', '\"')
        $cookie = $marshal.MarshalCookie.Replace('\', '\\').Replace('"', '\"')
        return "[${target}MarshalAs(UnmanagedType.CustomMarshaler, MarshalType = `"$marshalType`", MarshalCookie = `"$cookie`")] "
    }
    if ($marshal.Value -notin @(
            [Runtime.InteropServices.UnmanagedType]::BStr,
            [Runtime.InteropServices.UnmanagedType]::Interface,
            [Runtime.InteropServices.UnmanagedType]::Struct,
            [Runtime.InteropServices.UnmanagedType]::VariantBool,
            [Runtime.InteropServices.UnmanagedType]::IUnknown,
            [Runtime.InteropServices.UnmanagedType]::IDispatch
        )) { throw "Unsupported marshaling: $($marshal.Value)" }
    return "[${target}MarshalAs(UnmanagedType.$($marshal.Value))] "
}

function Get-ParameterDeclaration($Parameter, [switch]$Interop) {
    $attributes = ''
    if ($Interop) {
        if ($Parameter.IsIn) { $attributes += '[In] ' }
        if ($Parameter.IsOut) { $attributes += '[Out] ' }
        $attributes += Get-MarshalAttribute $Parameter
    }
    $direction = ''
    if ($Parameter.ParameterType.IsByRef) {
        $direction = if ($Parameter.IsOut -and -not $Parameter.IsIn) { 'out ' } else { 'ref ' }
    }
    return "$attributes$direction$(Get-TypeName $Parameter.ParameterType) @$($Parameter.Name)"
}

function Get-Argument($Parameter) {
    $direction = ''
    if ($Parameter.ParameterType.IsByRef) {
        $direction = if ($Parameter.IsOut -and -not $Parameter.IsIn) { 'out ' } else { 'ref ' }
    }
    return "$direction@$($Parameter.Name)"
}

$lines = [Collections.Generic.List[string]]::new()
$lines.Add('using System;')
$lines.Add('using System.Runtime.InteropServices;')
$lines.Add('')
$lines.Add('namespace OneNoteAutomation.Interop')
$lines.Add('{')
foreach ($enumType in ($types.Values | Where-Object IsEnum | Sort-Object Name)) {
    $lines.Add("    public enum $($enumType.Name) : $(Get-TypeName ([Enum]::GetUnderlyingType($enumType)))")
    $lines.Add('    {')
    foreach ($field in ($enumType.GetFields() | Where-Object IsLiteral | Sort-Object MetadataToken)) {
        $lines.Add("        @$($field.Name) = $($field.GetRawConstantValue()),")
    }
    $lines.Add('    }')
    $lines.Add('')
}

foreach ($structType in ($types.Values | Where-Object { $_.IsValueType -and -not $_.IsEnum } | Sort-Object Name)) {
    if (-not $structType.IsLayoutSequential) { throw "Unsupported structure layout: $structType" }
    $lines.Add("    [StructLayout(LayoutKind.Sequential, Pack = $($structType.StructLayoutAttribute.Pack), Size = $($structType.StructLayoutAttribute.Size))]")
    $lines.Add("    public struct $($structType.Name)")
    $lines.Add('    {')
    foreach ($field in ($structType.GetFields() | Sort-Object MetadataToken)) {
        $marshal = Get-MarshalAttribute $field
        $lines.Add("        ${marshal}public $(Get-TypeName $field.FieldType) @$($field.Name);")
    }
    $lines.Add('    }')
    $lines.Add('')
}

function Add-Members([type]$Type, [switch]$Proxy) {
    $properties = @($Type.GetProperties())
    $writtenProperties = [Collections.Generic.HashSet[string]]::new()
    foreach ($method in ($Type.GetMethods() | Sort-Object MetadataToken)) {
        $property = $properties | Where-Object {
            $_.GetGetMethod() -eq $method -or $_.GetSetMethod() -eq $method
        } | Select-Object -First 1
        if ($null -ne $property) {
            if (-not $writtenProperties.Add($property.Name)) { continue }
            $indexParameters = @($property.GetIndexParameters())
            $propertyName = "@$($property.Name)"
            if ($indexParameters.Count) {
                $indexDeclaration = ($indexParameters | ForEach-Object { Get-ParameterDeclaration $_ -Interop:(-not $Proxy) }) -join ', '
                $propertyName = "this[$indexDeclaration]"
            }
            $visibility = if ($Proxy) { 'public ' } else { '' }
            $lines.Add("        $visibility$(Get-TypeName $property.PropertyType) $propertyName")
            $lines.Add('        {')
            foreach ($accessor in @($property.GetGetMethod(), $property.GetSetMethod())) {
                if ($null -eq $accessor) { continue }
                $isGet = $accessor.Name.StartsWith('get_')
                $keyword = if ($isGet) { 'get' } else { 'set' }
                if ($Proxy) {
                    $target = "Application.@$($property.Name)"
                    if ($indexParameters.Count) {
                        $indexes = ($indexParameters | ForEach-Object { Get-Argument $_ }) -join ', '
                        $target = "Application[$indexes]"
                    }
                    $body = if ($isGet) { "return $target;" } else { "$target = value;" }
                    $lines.Add("            $keyword { $body }")
                }
                else {
                    $dispId = $accessor.GetCustomAttributes([Runtime.InteropServices.DispIdAttribute], $false) | Select-Object -First 1
                    if ($dispId) { $lines.Add("            [DispId($($dispId.Value))]") }
                    $returnMarshal = Get-MarshalAttribute $accessor.ReturnParameter -ReturnValue
                    if ($returnMarshal) { $lines.Add("            $returnMarshal") }
                    if (-not $isGet) {
                        $valueParameter = $accessor.GetParameters() | Select-Object -Last 1
                        if ($valueParameter.IsIn) { $lines.Add('            [param: In]') }
                        if ($valueParameter.IsOut) { $lines.Add('            [param: Out]') }
                        $valueMarshal = Get-MarshalAttribute $valueParameter
                        if ($valueMarshal) { $lines.Add('            ' + $valueMarshal.Replace('[MarshalAs', '[param: MarshalAs')) }
                    }
                    $lines.Add("            $keyword;")
                }
            }
            $lines.Add('        }')
        }
        else {
            $parameters = ($method.GetParameters() | ForEach-Object {
                    Get-ParameterDeclaration $_ -Interop:(-not $Proxy)
                }) -join ', '
            $signature = "$(Get-TypeName $method.ReturnType) @$($method.Name)($parameters)"
            if ($Proxy) {
                $arguments = ($method.GetParameters() | ForEach-Object { Get-Argument $_ }) -join ', '
                $returnKeyword = if ($method.ReturnType -eq [void]) { '' } else { 'return ' }
                $lines.Add("        public $signature")
                $lines.Add('        {')
                $lines.Add("            ${returnKeyword}Application.@$($method.Name)($arguments);")
                $lines.Add('        }')
            }
            else {
                $dispId = $method.GetCustomAttributes([Runtime.InteropServices.DispIdAttribute], $false) | Select-Object -First 1
                if ($dispId) { $lines.Add("        [DispId($($dispId.Value))]") }
                $returnMarshal = Get-MarshalAttribute $method.ReturnParameter -ReturnValue
                if ($returnMarshal) { $lines.Add("        $returnMarshal") }
                $lines.Add("        $signature;")
            }
        }
        $lines.Add('')
    }
}

foreach ($interfaceType in ($types.Values | Where-Object IsInterface | Sort-Object Name)) {
    $interfaceKind = $interfaceType.GetCustomAttributes([Runtime.InteropServices.InterfaceTypeAttribute], $false) | Select-Object -First 1
    $kind = if ($interfaceKind) { $interfaceKind.Value.ToString() } else { 'InterfaceIsDual' }
    $lines.Add('    [ComImport]')
    $lines.Add("    [Guid(`"$($interfaceType.GUID)`")]")
    $lines.Add("    [InterfaceType(ComInterfaceType.$kind)]")
    $lines.Add("    public interface $(Get-TypeName $interfaceType)")
    $lines.Add('    {')
    Add-Members $interfaceType
    $lines.Add('    }')
    $lines.Add('')
}
function Get-OverloadArguments([Reflection.MethodInfo]$Method) {
    $body = $Method.GetMethodBody().GetILAsByteArray()
    $stack = [Collections.Generic.List[object]]::new()
    $parameters = $Method.GetParameters()
    $offset = 0
    while ($offset -lt $body.Length) {
        $opcode = $body[$offset++]
        if ($opcode -eq 0x02) {
            $stack.Add('this')
        }
        elseif ($opcode -in 0x03, 0x04, 0x05, 0x0E) {
            $argumentIndex = if ($opcode -eq 0x0E) { [int]$body[$offset++] } else { $opcode - 0x02 }
            if ($argumentIndex -eq 0) { throw "Unexpected this argument in $Method" }
            $stack.Add((Get-Argument $parameters[$argumentIndex - 1]))
        }
        elseif ($opcode -ge 0x15 -and $opcode -le 0x1E) {
            $stack.Add([int]($opcode - 0x16))
        }
        elseif ($opcode -eq 0x20) {
            $stack.Add([BitConverter]::ToInt32($body, $offset))
            $offset += 4
        }
        elseif ($opcode -eq 0x72) {
            $text = $Method.Module.ResolveString([BitConverter]::ToInt32($body, $offset))
            $stack.Add(('"' + $text.Replace('\', '\\').Replace('"', '\"').Replace("`r", '\r').Replace("`n", '\n') + '"'))
            $offset += 4
        }
        elseif ($opcode -eq 0x7E) {
            $field = $Method.Module.ResolveField([BitConverter]::ToInt32($body, $offset))
            if ($field.DeclaringType -ne [datetime] -or $field.Name -ne 'MinValue') {
                throw "Unsupported default field: $field"
            }
            $stack.Add('System.DateTime.MinValue')
            $offset += 4
        }
        elseif ($opcode -in 0x28, 0x6F) {
            $target = $Method.Module.ResolveMethod([BitConverter]::ToInt32($body, $offset))
            $offset += 4
            $targetParameters = $target.GetParameters()
            if ($target.Name -notin $Method.Name, ($Method.Name + '_AutoBusyRetry') -or $target.IsStatic -or
                $targetParameters.Count -le $parameters.Count -or
                $stack.Count -ne ($targetParameters.Count + 1) -or $stack[0] -ne 'this' -or
                $offset -ne ($body.Length - 1) -or $body[$offset] -ne 0x2A) {
                throw "Unsupported overload forwarding body: $Method"
            }
            $arguments = for ($index = 0; $index -lt $targetParameters.Count; $index++) {
                $value = $stack[$index + 1]
                $targetType = $targetParameters[$index].ParameterType
                if ($value -is [int]) {
                    if ($targetType.IsEnum) {
                        $enumName = [Enum]::GetName($targetType, $value)
                        if ($null -eq $enumName) { throw "Unknown enum default: $targetType $value" }
                        "$(Get-TypeName $targetType).@$enumName"
                    }
                    elseif ($targetType -eq [bool]) {
                        if ($value -notin 0, 1) { throw "Invalid Boolean default: $value" }
                        if ($value) { 'true' } else { 'false' }
                    }
                    elseif ($targetType -eq [int]) { "$value" }
                    else { throw "Unsupported integer default for $targetType" }
                }
                else { $value }
            }
            return $arguments -join ', '
        }
        else { throw "Unsupported IL opcode $opcode in $Method" }
    }
    throw "No forwarding call found in $Method"
}

$lines.Add('    public sealed partial class OneNoteApplication : IOneNoteApplication, IDisposable')
$lines.Add('    {')
Add-Members $applicationType -Proxy
$overloadCount = 0
foreach ($method in ($managedInterface.GetMethods() | Where-Object { -not $_.IsSpecialName } | Sort-Object MetadataToken)) {
    $parameterTypes = [type[]]@($method.GetParameters() | ForEach-Object ParameterType)
    if ($null -ne $applicationType.GetMethod($method.Name, $parameterTypes)) { continue }
    $implementation = $managedClass.GetMethod($method.Name,
        [Reflection.BindingFlags]'Public,Instance,DeclaredOnly', $null, $parameterTypes, $null)
    if ($null -eq $implementation) { throw "Managed overload implementation not found: $method" }
    $arguments = Get-OverloadArguments $implementation
    $parameters = ($method.GetParameters() | ForEach-Object { Get-ParameterDeclaration $_ }) -join ', '
    $lines.Add("        public $(Get-TypeName $method.ReturnType) @$($method.Name)($parameters)")
    $lines.Add('        {')
    $lines.Add("            @$($method.Name)($arguments);")
    $lines.Add('        }')
    $lines.Add('')
    $overloadCount++
}
$lines.Add('    }')
$lines.Add('}')
$outputPath = Join-Path $PSScriptRoot '../OneNoteAutomation/Interop/OneNoteInterop.Generated.cs'
[IO.File]::WriteAllLines($outputPath, $lines, [Text.UTF8Encoding]::new($false))
[pscustomobject]@{
    SourceAssembly         = $assembly.FullName
    GeneratedPath          = $outputPath
    InterfaceCount         = @($types.Values | Where-Object IsInterface).Count
    EnumCount              = @($types.Values | Where-Object IsEnum).Count
    ApplicationMethodCount = $applicationType.GetMethods().Count
    ConvenienceOverloads   = $overloadCount
}