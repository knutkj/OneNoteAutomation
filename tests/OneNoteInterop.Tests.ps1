Describe 'OneNote managed interop' -Skip:($PSEdition -ne 'Core') {
    BeforeAll {
        Import-Module (Join-Path $PSScriptRoot '../OneNoteAutomation/OneNoteAutomation.psd1') -Force
        & (Get-Module OneNoteAutomation) { Initialize-OneNoteInterop }

        if (-not ('OneNoteRecordingProxy' -as [type])) {
            Add-Type -TypeDefinition @'
using System;
using System.Reflection;

public class OneNoteRecordingProxy : DispatchProxy
{
    public string LastMethod;
    public object[] LastArguments;
    public string Output = "<Page ID='test-page-id' />";

    protected override object Invoke(MethodInfo method, object[] arguments)
    {
        LastMethod = method.Name;
        LastArguments = (object[])arguments.Clone();
        ParameterInfo[] parameters = method.GetParameters();
        for (int index = 0; index < parameters.Length; index++)
        {
            if (parameters[index].IsOut) arguments[index] = Output;
        }
        return null;
    }
}
'@
        }

        $createProxy = [System.Reflection.DispatchProxy].GetMethods() |
            Where-Object { $_.Name -eq 'Create' -and $_.IsGenericMethodDefinition }
        $createProxy = $createProxy.MakeGenericMethod(
            [OneNoteAutomation.Interop.IOneNoteApplication], [OneNoteRecordingProxy])
        $constructor = [OneNoteAutomation.Interop.OneNoteApplication].GetConstructors(
            [System.Reflection.BindingFlags]'Instance, NonPublic')[0]
    }

    BeforeEach {
        $native = $createProxy.Invoke($null, @())
        $releasedObjects = [System.Collections.Generic.List[object]]::new()
        $release = [System.Action[object]]{
            param($instance)
            $releasedObjects.Add($instance)
        }.GetNewClosure()
        $application = $constructor.Invoke([object[]]@($native, $release))
    }

    AfterEach {
        if ($application) { $application.Dispose() }
    }

    It 'forwards hierarchy arguments and returns XML through a reference' {
        [string]$xml = ''
        $application.GetHierarchy('test-section-id', 4, [ref]$xml)

        $xml | Should -Be $native.Output
        $native.LastMethod | Should -Be 'GetHierarchy'
        $native.LastArguments[0] | Should -Be 'test-section-id'
        [int]$native.LastArguments[1] | Should -Be 4
        [int]$native.LastArguments[3] | Should -Be 2
    }

    It 'forwards page-content and page-creation reference parameters' {
        [string]$xml = ''
        $application.GetPageContent('test-page-id', [ref]$xml)
        $xml | Should -Be $native.Output
        $native.LastMethod | Should -Be 'GetPageContent'
        [int]$native.LastArguments[2] | Should -Be 0
        [int]$native.LastArguments[3] | Should -Be 2

        $native.Output = 'new-page-id'
        [string]$pageId = ''
        $application.CreateNewPage('test-section-id', [ref]$pageId, 1)
        $pageId | Should -Be 'new-page-id'
        $native.LastMethod | Should -Be 'CreateNewPage'
        $native.LastArguments[0] | Should -Be 'test-section-id'
        [int]$native.LastArguments[2] | Should -Be 1
    }

    It 'supplies the existing update and navigation defaults' {
        $application.UpdateHierarchy('<Section />')
        $native.LastMethod | Should -Be 'UpdateHierarchy'
        [int]$native.LastArguments[1] | Should -Be 2

        $application.UpdatePageContent('<Page />')
        $native.LastMethod | Should -Be 'UpdatePageContent'
        $native.LastArguments[0] | Should -Be '<Page />'
        $native.LastArguments[1] | Should -Be ([datetime]::MinValue)
        [int]$native.LastArguments[2] | Should -Be 2
        $native.LastArguments[3] | Should -BeFalse

        $application.NavigateTo('test-page-id')
        $native.LastMethod | Should -Be 'NavigateTo'
        $native.LastArguments[0] | Should -Be 'test-page-id'
        $native.LastArguments[1] | Should -Be ''
        $native.LastArguments[2] | Should -BeFalse
    }

    It 'releases exactly once and rejects calls after disposal' {
        Remove-ComObject -ComObject $application
        $application.Dispose()

        $releasedObjects.Count | Should -Be 1
        [object]::ReferenceEquals($releasedObjects[0], $native) | Should -BeTrue
        { $application.NavigateTo('test-page-id') } |
            Should -Throw -ExpectedMessage '*disposed*'
    }
}