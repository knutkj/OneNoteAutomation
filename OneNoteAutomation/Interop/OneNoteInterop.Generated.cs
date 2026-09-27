using System;
using System.Runtime.InteropServices;

namespace OneNoteAutomation.Interop
{
    public enum CreateFileType : int
    {
        @cftNone = 0,
        @cftNotebook = 1,
        @cftFolder = 2,
        @cftSection = 3,
    }

    public enum DockLocation : int
    {
        @dlDefault = -1,
        @dlNone = 0,
        @dlLeft = 1,
        @dlRight = 2,
        @dlTop = 3,
        @dlBottom = 4,
    }

    public enum FilingLocation : int
    {
        @flEMail = 0,
        @flContacts = 1,
        @flTasks = 2,
        @flMeetings = 3,
        @flWebContent = 4,
        @flPrintOuts = 5,
    }

    public enum FilingLocationType : int
    {
        @fltNamedSectionNewPage = 0,
        @fltCurrentSectionNewPage = 1,
        @fltCurrentPage = 2,
        @fltNamedPage = 4,
    }

    public enum HierarchyElement : int
    {
        @heNone = 0,
        @heNotebooks = 1,
        @heSectionGroups = 2,
        @heSections = 4,
        @hePages = 8,
    }

    public enum HierarchyScope : int
    {
        @hsSelf = 0,
        @hsChildren = 1,
        @hsNotebooks = 2,
        @hsSections = 3,
        @hsPages = 4,
    }

    public enum NewPageStyle : int
    {
        @npsDefault = 0,
        @npsBlankPageWithTitle = 1,
        @npsBlankPageNoTitle = 2,
    }

    public enum NotebookFilterOutType : int
    {
        @nfoLocal = 1,
        @nfoNetwork = 2,
        @nfoWeb = 4,
        @nfoNoWacUrl = 8,
    }

    public enum PageInfo : int
    {
        @piBasic = 0,
        @piBinaryData = 1,
        @piSelection = 2,
        @piFileType = 4,
        @piBinaryDataSelection = 3,
        @piBinaryDataFileType = 5,
        @piSelectionFileType = 6,
        @piAll = 7,
    }

    public enum PublishFormat : int
    {
        @pfOneNote = 0,
        @pfOneNotePackage = 1,
        @pfMHTML = 2,
        @pfPDF = 3,
        @pfXPS = 4,
        @pfWord = 5,
        @pfEMF = 6,
        @pfHTML = 7,
        @pfOneNote2007 = 8,
    }

    public enum RecentResultType : int
    {
        @rrtNone = 0,
        @rrtFiling = 1,
        @rrtSearch = 2,
        @rrtLinks = 3,
    }

    public enum SpecialLocation : int
    {
        @slBackUpFolder = 0,
        @slUnfiledNotesSection = 1,
        @slDefaultNotebookFolder = 2,
    }

    public enum TreeCollapsedStateType : int
    {
        @tcsExpanded = 0,
        @tcsCollapsed = 1,
    }

    public enum XMLSchema : int
    {
        @xs2007 = 0,
        @xs2010 = 1,
        @xs2013 = 2,
        @xsCurrent = 2,
    }

    [StructLayout(LayoutKind.Sequential, Pack = 4, Size = 0)]
    public struct tagPOINT
    {
        public int @x;
        public int @y;
    }

    [ComImport]
    [Guid("452ac71a-b655-4967-a208-a4cc39dd7949")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IOneNoteApplication
    {
        [DispId(1610743808)]
        void @GetHierarchy([In] [MarshalAs(UnmanagedType.BStr)] string @bstrStartNodeID, [In] HierarchyScope @hsScope, [Out] [MarshalAs(UnmanagedType.BStr)] out string @pbstrHierarchyXmlOut, [In] XMLSchema @xsSchema);

        [DispId(1610743809)]
        void @UpdateHierarchy([In] [MarshalAs(UnmanagedType.BStr)] string @bstrChangesXmlIn, [In] XMLSchema @xsSchema);

        [DispId(1610743810)]
        void @OpenHierarchy([In] [MarshalAs(UnmanagedType.BStr)] string @bstrPath, [In] [MarshalAs(UnmanagedType.BStr)] string @bstrRelativeToObjectID, [Out] [MarshalAs(UnmanagedType.BStr)] out string @pbstrObjectID, [In] CreateFileType @cftIfNotExist);

        [DispId(1610743811)]
        void @DeleteHierarchy([In] [MarshalAs(UnmanagedType.BStr)] string @bstrObjectID, [In] System.DateTime @dateExpectedLastModified, [In] bool @deletePermanently);

        [DispId(1610743812)]
        void @CreateNewPage([In] [MarshalAs(UnmanagedType.BStr)] string @bstrSectionID, [Out] [MarshalAs(UnmanagedType.BStr)] out string @pbstrPageID, [In] NewPageStyle @npsNewPageStyle);

        [DispId(1610743813)]
        void @CloseNotebook([In] [MarshalAs(UnmanagedType.BStr)] string @bstrNotebookID, [In] bool @force);

        [DispId(1610743814)]
        void @GetHierarchyParent([In] [MarshalAs(UnmanagedType.BStr)] string @bstrObjectID, [Out] [MarshalAs(UnmanagedType.BStr)] out string @pbstrParentID);

        [DispId(1610743815)]
        void @GetPageContent([In] [MarshalAs(UnmanagedType.BStr)] string @bstrPageID, [Out] [MarshalAs(UnmanagedType.BStr)] out string @pbstrPageXmlOut, [In] PageInfo @pageInfoToExport, [In] XMLSchema @xsSchema);

        [DispId(1610743816)]
        void @UpdatePageContent([In] [MarshalAs(UnmanagedType.BStr)] string @bstrPageChangesXmlIn, [In] System.DateTime @dateExpectedLastModified, [In] XMLSchema @xsSchema, [In] bool @force);

        [DispId(1610743817)]
        void @GetBinaryPageContent([In] [MarshalAs(UnmanagedType.BStr)] string @bstrPageID, [In] [MarshalAs(UnmanagedType.BStr)] string @bstrCallbackID, [Out] [MarshalAs(UnmanagedType.BStr)] out string @pbstrBinaryObjectB64Out);

        [DispId(1610743818)]
        void @DeletePageContent([In] [MarshalAs(UnmanagedType.BStr)] string @bstrPageID, [In] [MarshalAs(UnmanagedType.BStr)] string @bstrObjectID, [In] System.DateTime @dateExpectedLastModified, [In] bool @force);

        [DispId(1610743819)]
        void @NavigateTo([In] [MarshalAs(UnmanagedType.BStr)] string @bstrHierarchyObjectID, [In] [MarshalAs(UnmanagedType.BStr)] string @bstrObjectID, [In] bool @fNewWindow);

        [DispId(1610743820)]
        void @NavigateToUrl([In] [MarshalAs(UnmanagedType.BStr)] string @bstrUrl, [In] bool @fNewWindow);

        [DispId(1610743821)]
        void @Publish([In] [MarshalAs(UnmanagedType.BStr)] string @bstrHierarchyID, [In] [MarshalAs(UnmanagedType.BStr)] string @bstrTargetFilePath, [In] PublishFormat @pfPublishFormat, [In] [MarshalAs(UnmanagedType.BStr)] string @bstrCLSIDofExporter);

        [DispId(1610743822)]
        void @OpenPackage([In] [MarshalAs(UnmanagedType.BStr)] string @bstrPathPackage, [In] [MarshalAs(UnmanagedType.BStr)] string @bstrPathDest, [Out] [MarshalAs(UnmanagedType.BStr)] out string @pbstrPathOut);

        [DispId(1610743823)]
        void @GetHyperlinkToObject([In] [MarshalAs(UnmanagedType.BStr)] string @bstrHierarchyID, [In] [MarshalAs(UnmanagedType.BStr)] string @bstrPageContentObjectID, [Out] [MarshalAs(UnmanagedType.BStr)] out string @pbstrHyperlinkOut);

        [DispId(1610743824)]
        void @FindPages([In] [MarshalAs(UnmanagedType.BStr)] string @bstrStartNodeID, [In] [MarshalAs(UnmanagedType.BStr)] string @bstrSearchString, [Out] [MarshalAs(UnmanagedType.BStr)] out string @pbstrHierarchyXmlOut, [In] bool @fIncludeUnindexedPages, [In] bool @fDisplay, [In] XMLSchema @xsSchema);

        [DispId(1610743825)]
        void @FindMeta([In] [MarshalAs(UnmanagedType.BStr)] string @bstrStartNodeID, [In] [MarshalAs(UnmanagedType.BStr)] string @bstrSearchStringName, [Out] [MarshalAs(UnmanagedType.BStr)] out string @pbstrHierarchyXmlOut, [In] bool @fIncludeUnindexedPages, [In] XMLSchema @xsSchema);

        [DispId(1610743826)]
        void @GetSpecialLocation([In] SpecialLocation @slToGet, [Out] [MarshalAs(UnmanagedType.BStr)] out string @pbstrSpecialLocationPath);

        [DispId(1610743827)]
        void @MergeFiles([In] [MarshalAs(UnmanagedType.BStr)] string @bstrBaseFile, [In] [MarshalAs(UnmanagedType.BStr)] string @bstrClientFile, [In] [MarshalAs(UnmanagedType.BStr)] string @bstrServerFile, [In] [MarshalAs(UnmanagedType.BStr)] string @bstrTargetFile);

        [DispId(1610743828)]
        [return: MarshalAs(UnmanagedType.Interface)] 
        IQuickFilingDialog @QuickFiling();

        [DispId(1610743829)]
        void @SyncHierarchy([In] [MarshalAs(UnmanagedType.BStr)] string @bstrHierarchyID);

        [DispId(1610743830)]
        void @SetFilingLocation([In] FilingLocation @flToSet, [In] FilingLocationType @fltToSet, [In] [MarshalAs(UnmanagedType.BStr)] string @bstrFilingSectionID);

        Windows @Windows
        {
            [DispId(100)]
            [return: MarshalAs(UnmanagedType.Interface)] 
            get;
        }

        bool @Dummy1
        {
            [DispId(102)]
            get;
        }

        [DispId(1610743833)]
        void @MergeSections([In] [MarshalAs(UnmanagedType.BStr)] string @bstrSectionSourceId, [In] [MarshalAs(UnmanagedType.BStr)] string @bstrSectionDestinationId);

        object @COMAddIns
        {
            [DispId(104)]
            [return: MarshalAs(UnmanagedType.IDispatch)] 
            get;
        }

        object @LanguageSettings
        {
            [DispId(105)]
            [return: MarshalAs(UnmanagedType.IDispatch)] 
            get;
        }

        [DispId(1610743836)]
        void @GetWebHyperlinkToObject([In] [MarshalAs(UnmanagedType.BStr)] string @bstrHierarchyID, [In] [MarshalAs(UnmanagedType.BStr)] string @bstrPageContentObjectID, [Out] [MarshalAs(UnmanagedType.BStr)] out string @pbstrHyperlinkOut);

    }

    [ComImport]
    [Guid("1d12bd3f-89b6-4077-aa2c-c9dc2bca42f9")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IQuickFilingDialog
    {
        string @Title
        {
            [DispId(0)]
            [return: MarshalAs(UnmanagedType.BStr)] 
            get;
            [DispId(0)]
            [param: In]
            [param: MarshalAs(UnmanagedType.BStr)] 
            set;
        }

        string @Description
        {
            [DispId(1)]
            [return: MarshalAs(UnmanagedType.BStr)] 
            get;
            [DispId(1)]
            [param: In]
            [param: MarshalAs(UnmanagedType.BStr)] 
            set;
        }

        string @CheckboxText
        {
            [DispId(2)]
            [return: MarshalAs(UnmanagedType.BStr)] 
            get;
            [DispId(2)]
            [param: In]
            [param: MarshalAs(UnmanagedType.BStr)] 
            set;
        }

        bool @CheckboxState
        {
            [DispId(3)]
            get;
            [DispId(3)]
            [param: In]
            set;
        }

        ulong @WindowHandle
        {
            [DispId(4)]
            get;
        }

        HierarchyElement @TreeDepth
        {
            [DispId(5)]
            get;
            [DispId(5)]
            [param: In]
            set;
        }

        ulong @ParentWindowHandle
        {
            [DispId(6)]
            get;
            [DispId(6)]
            [param: In]
            set;
        }

        tagPOINT @Position
        {
            [DispId(7)]
            get;
            [DispId(7)]
            [param: In]
            set;
        }

        [DispId(8)]
        void @SetRecentResults([In] RecentResultType @recentResults, [In] bool @fShowCurrentSection, [In] bool @fShowCurrentPage, [In] bool @fShowUnfiledNotes);

        [DispId(10)]
        void @AddButton([In] [MarshalAs(UnmanagedType.BStr)] string @bstrText, [In] HierarchyElement @allowedElements, [In] HierarchyElement @allowedReadOnlyElements, [In] bool @fDefault);

        [DispId(11)]
        void @Run([In] [MarshalAs(UnmanagedType.Interface)] IQuickFilingDialogCallback @piCallback);

        string @SelectedItem
        {
            [DispId(12)]
            [return: MarshalAs(UnmanagedType.BStr)] 
            get;
        }

        uint @PressedButton
        {
            [DispId(13)]
            get;
        }

        TreeCollapsedStateType @TreeCollapsedState
        {
            [DispId(14)]
            [param: In]
            set;
        }

        NotebookFilterOutType @NotebookFilterOut
        {
            [DispId(15)]
            [param: In]
            set;
        }

        [DispId(16)]
        void @ShowCreateNewNotebook();

        [DispId(17)]
        void @AddInitialEditor([MarshalAs(UnmanagedType.BStr)] string @initialEditor);

        [DispId(18)]
        void @ClearInitialEditors();

        [DispId(19)]
        void @ShowSharingHyperlink();

    }

    [ComImport]
    [Guid("627ea7b4-95b5-4980-84c1-9d20da4460b1")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IQuickFilingDialogCallback
    {
        [DispId(1610743808)]
        void @OnDialogClosed([In] [MarshalAs(UnmanagedType.Interface)] IQuickFilingDialog @dialog);

    }

    [ComImport]
    [Guid("8e8304b8-cbd1-44f8-b0e8-89c625b2002e")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface Window
    {
        ulong @WindowHandle
        {
            [DispId(0)]
            get;
        }

        string @CurrentPageId
        {
            [DispId(1)]
            [return: MarshalAs(UnmanagedType.BStr)] 
            get;
        }

        string @CurrentSectionId
        {
            [DispId(2)]
            [return: MarshalAs(UnmanagedType.BStr)] 
            get;
        }

        string @CurrentSectionGroupId
        {
            [DispId(3)]
            [return: MarshalAs(UnmanagedType.BStr)] 
            get;
        }

        string @CurrentNotebookId
        {
            [DispId(4)]
            [return: MarshalAs(UnmanagedType.BStr)] 
            get;
        }

        [DispId(9)]
        void @NavigateTo([In] [MarshalAs(UnmanagedType.BStr)] string @bstrHierarchyObjectID, [In] [MarshalAs(UnmanagedType.BStr)] string @bstrObjectID);

        bool @FullPageView
        {
            [DispId(10)]
            get;
            [DispId(10)]
            set;
        }

        bool @Active
        {
            [DispId(11)]
            get;
            [DispId(11)]
            set;
        }

        DockLocation @DockedLocation
        {
            [DispId(13)]
            get;
            [DispId(13)]
            set;
        }

        IOneNoteApplication @Application
        {
            [DispId(14)]
            [return: MarshalAs(UnmanagedType.Interface)] 
            get;
        }

        bool @SideNote
        {
            [DispId(15)]
            get;
        }

        [DispId(16)]
        void @NavigateToUrl([In] [MarshalAs(UnmanagedType.BStr)] string @bstrUrl);

        [DispId(17)]
        void @SetDockedLocation([In] DockLocation @DockLocation, [In] tagPOINT @ptMonitor);

    }

    [ComImport]
    [Guid("6d4b9c3e-cc05-493f-85e2-43d1006df96a")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface Windows
    {
        Window this[[In] uint @Index]
        {
            [DispId(0)]
            [return: MarshalAs(UnmanagedType.Interface)] 
            get;
        }

        uint @Count
        {
            [DispId(1)]
            get;
        }

        [DispId(-4)]
        [return: MarshalAs(UnmanagedType.CustomMarshaler, MarshalType = "System.Runtime.InteropServices.CustomMarshalers.EnumeratorToEnumVariantMarshaler, CustomMarshalers, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", MarshalCookie = "")] 
        System.Collections.IEnumerator @GetEnumerator();

        Window @CurrentWindow
        {
            [DispId(3)]
            [return: MarshalAs(UnmanagedType.Interface)] 
            get;
        }

    }

    public sealed partial class OneNoteApplication : IOneNoteApplication, IDisposable
    {
        public void @GetHierarchy(string @bstrStartNodeID, HierarchyScope @hsScope, out string @pbstrHierarchyXmlOut, XMLSchema @xsSchema)
        {
            Application.@GetHierarchy(@bstrStartNodeID, @hsScope, out @pbstrHierarchyXmlOut, @xsSchema);
        }

        public void @UpdateHierarchy(string @bstrChangesXmlIn, XMLSchema @xsSchema)
        {
            Application.@UpdateHierarchy(@bstrChangesXmlIn, @xsSchema);
        }

        public void @OpenHierarchy(string @bstrPath, string @bstrRelativeToObjectID, out string @pbstrObjectID, CreateFileType @cftIfNotExist)
        {
            Application.@OpenHierarchy(@bstrPath, @bstrRelativeToObjectID, out @pbstrObjectID, @cftIfNotExist);
        }

        public void @DeleteHierarchy(string @bstrObjectID, System.DateTime @dateExpectedLastModified, bool @deletePermanently)
        {
            Application.@DeleteHierarchy(@bstrObjectID, @dateExpectedLastModified, @deletePermanently);
        }

        public void @CreateNewPage(string @bstrSectionID, out string @pbstrPageID, NewPageStyle @npsNewPageStyle)
        {
            Application.@CreateNewPage(@bstrSectionID, out @pbstrPageID, @npsNewPageStyle);
        }

        public void @CloseNotebook(string @bstrNotebookID, bool @force)
        {
            Application.@CloseNotebook(@bstrNotebookID, @force);
        }

        public void @GetHierarchyParent(string @bstrObjectID, out string @pbstrParentID)
        {
            Application.@GetHierarchyParent(@bstrObjectID, out @pbstrParentID);
        }

        public void @GetPageContent(string @bstrPageID, out string @pbstrPageXmlOut, PageInfo @pageInfoToExport, XMLSchema @xsSchema)
        {
            Application.@GetPageContent(@bstrPageID, out @pbstrPageXmlOut, @pageInfoToExport, @xsSchema);
        }

        public void @UpdatePageContent(string @bstrPageChangesXmlIn, System.DateTime @dateExpectedLastModified, XMLSchema @xsSchema, bool @force)
        {
            Application.@UpdatePageContent(@bstrPageChangesXmlIn, @dateExpectedLastModified, @xsSchema, @force);
        }

        public void @GetBinaryPageContent(string @bstrPageID, string @bstrCallbackID, out string @pbstrBinaryObjectB64Out)
        {
            Application.@GetBinaryPageContent(@bstrPageID, @bstrCallbackID, out @pbstrBinaryObjectB64Out);
        }

        public void @DeletePageContent(string @bstrPageID, string @bstrObjectID, System.DateTime @dateExpectedLastModified, bool @force)
        {
            Application.@DeletePageContent(@bstrPageID, @bstrObjectID, @dateExpectedLastModified, @force);
        }

        public void @NavigateTo(string @bstrHierarchyObjectID, string @bstrObjectID, bool @fNewWindow)
        {
            Application.@NavigateTo(@bstrHierarchyObjectID, @bstrObjectID, @fNewWindow);
        }

        public void @NavigateToUrl(string @bstrUrl, bool @fNewWindow)
        {
            Application.@NavigateToUrl(@bstrUrl, @fNewWindow);
        }

        public void @Publish(string @bstrHierarchyID, string @bstrTargetFilePath, PublishFormat @pfPublishFormat, string @bstrCLSIDofExporter)
        {
            Application.@Publish(@bstrHierarchyID, @bstrTargetFilePath, @pfPublishFormat, @bstrCLSIDofExporter);
        }

        public void @OpenPackage(string @bstrPathPackage, string @bstrPathDest, out string @pbstrPathOut)
        {
            Application.@OpenPackage(@bstrPathPackage, @bstrPathDest, out @pbstrPathOut);
        }

        public void @GetHyperlinkToObject(string @bstrHierarchyID, string @bstrPageContentObjectID, out string @pbstrHyperlinkOut)
        {
            Application.@GetHyperlinkToObject(@bstrHierarchyID, @bstrPageContentObjectID, out @pbstrHyperlinkOut);
        }

        public void @FindPages(string @bstrStartNodeID, string @bstrSearchString, out string @pbstrHierarchyXmlOut, bool @fIncludeUnindexedPages, bool @fDisplay, XMLSchema @xsSchema)
        {
            Application.@FindPages(@bstrStartNodeID, @bstrSearchString, out @pbstrHierarchyXmlOut, @fIncludeUnindexedPages, @fDisplay, @xsSchema);
        }

        public void @FindMeta(string @bstrStartNodeID, string @bstrSearchStringName, out string @pbstrHierarchyXmlOut, bool @fIncludeUnindexedPages, XMLSchema @xsSchema)
        {
            Application.@FindMeta(@bstrStartNodeID, @bstrSearchStringName, out @pbstrHierarchyXmlOut, @fIncludeUnindexedPages, @xsSchema);
        }

        public void @GetSpecialLocation(SpecialLocation @slToGet, out string @pbstrSpecialLocationPath)
        {
            Application.@GetSpecialLocation(@slToGet, out @pbstrSpecialLocationPath);
        }

        public void @MergeFiles(string @bstrBaseFile, string @bstrClientFile, string @bstrServerFile, string @bstrTargetFile)
        {
            Application.@MergeFiles(@bstrBaseFile, @bstrClientFile, @bstrServerFile, @bstrTargetFile);
        }

        public IQuickFilingDialog @QuickFiling()
        {
            return Application.@QuickFiling();
        }

        public void @SyncHierarchy(string @bstrHierarchyID)
        {
            Application.@SyncHierarchy(@bstrHierarchyID);
        }

        public void @SetFilingLocation(FilingLocation @flToSet, FilingLocationType @fltToSet, string @bstrFilingSectionID)
        {
            Application.@SetFilingLocation(@flToSet, @fltToSet, @bstrFilingSectionID);
        }

        public Windows @Windows
        {
            get { return Application.@Windows; }
        }

        public bool @Dummy1
        {
            get { return Application.@Dummy1; }
        }

        public void @MergeSections(string @bstrSectionSourceId, string @bstrSectionDestinationId)
        {
            Application.@MergeSections(@bstrSectionSourceId, @bstrSectionDestinationId);
        }

        public object @COMAddIns
        {
            get { return Application.@COMAddIns; }
        }

        public object @LanguageSettings
        {
            get { return Application.@LanguageSettings; }
        }

        public void @GetWebHyperlinkToObject(string @bstrHierarchyID, string @bstrPageContentObjectID, out string @pbstrHyperlinkOut)
        {
            Application.@GetWebHyperlinkToObject(@bstrHierarchyID, @bstrPageContentObjectID, out @pbstrHyperlinkOut);
        }

        public void @GetHierarchy(string @bstrStartNodeID, HierarchyScope @hsScope, out string @pbstrHierarchyXmlOut)
        {
            @GetHierarchy(@bstrStartNodeID, @hsScope, out @pbstrHierarchyXmlOut, XMLSchema.@xs2013);
        }

        public void @UpdateHierarchy(string @bstrChangesXmlIn)
        {
            @UpdateHierarchy(@bstrChangesXmlIn, XMLSchema.@xs2013);
        }

        public void @OpenHierarchy(string @bstrPath, string @bstrRelativeToObjectID, out string @pbstrObjectID)
        {
            @OpenHierarchy(@bstrPath, @bstrRelativeToObjectID, out @pbstrObjectID, CreateFileType.@cftNone);
        }

        public void @DeleteHierarchy(string @bstrObjectID, System.DateTime @dateExpectedLastModified)
        {
            @DeleteHierarchy(@bstrObjectID, @dateExpectedLastModified, false);
        }

        public void @DeleteHierarchy(string @bstrObjectID)
        {
            @DeleteHierarchy(@bstrObjectID, System.DateTime.MinValue, false);
        }

        public void @CreateNewPage(string @bstrSectionID, out string @pbstrPageID)
        {
            @CreateNewPage(@bstrSectionID, out @pbstrPageID, NewPageStyle.@npsDefault);
        }

        public void @CloseNotebook(string @bstrNotebookID)
        {
            @CloseNotebook(@bstrNotebookID, false);
        }

        public void @GetPageContent(string @bstrPageID, out string @pbstrPageXmlOut, PageInfo @pageInfoToExport)
        {
            @GetPageContent(@bstrPageID, out @pbstrPageXmlOut, @pageInfoToExport, XMLSchema.@xs2013);
        }

        public void @GetPageContent(string @bstrPageID, out string @pbstrPageXmlOut)
        {
            @GetPageContent(@bstrPageID, out @pbstrPageXmlOut, PageInfo.@piBasic, XMLSchema.@xs2013);
        }

        public void @UpdatePageContent(string @bstrPageChangesXmlIn, System.DateTime @dateExpectedLastModified, XMLSchema @xsSchema)
        {
            @UpdatePageContent(@bstrPageChangesXmlIn, @dateExpectedLastModified, @xsSchema, false);
        }

        public void @UpdatePageContent(string @bstrPageChangesXmlIn, System.DateTime @dateExpectedLastModified)
        {
            @UpdatePageContent(@bstrPageChangesXmlIn, @dateExpectedLastModified, XMLSchema.@xs2013, false);
        }

        public void @UpdatePageContent(string @bstrPageChangesXmlIn)
        {
            @UpdatePageContent(@bstrPageChangesXmlIn, System.DateTime.MinValue, XMLSchema.@xs2013, false);
        }

        public void @DeletePageContent(string @bstrPageID, string @bstrObjectID, System.DateTime @dateExpectedLastModified)
        {
            @DeletePageContent(@bstrPageID, @bstrObjectID, @dateExpectedLastModified, false);
        }

        public void @DeletePageContent(string @bstrPageID, string @bstrObjectID)
        {
            @DeletePageContent(@bstrPageID, @bstrObjectID, System.DateTime.MinValue, false);
        }

        public void @NavigateTo(string @bstrHierarchyObjectID, string @bstrObjectID)
        {
            @NavigateTo(@bstrHierarchyObjectID, @bstrObjectID, false);
        }

        public void @NavigateTo(string @bstrHierarchyObjectID)
        {
            @NavigateTo(@bstrHierarchyObjectID, "", false);
        }

        public void @NavigateToUrl(string @bstrUrl)
        {
            @NavigateToUrl(@bstrUrl, false);
        }

        public void @Publish(string @bstrHierarchyID, string @bstrTargetFilePath, PublishFormat @pfPublishFormat)
        {
            @Publish(@bstrHierarchyID, @bstrTargetFilePath, @pfPublishFormat, "");
        }

        public void @Publish(string @bstrHierarchyID, string @bstrTargetFilePath)
        {
            @Publish(@bstrHierarchyID, @bstrTargetFilePath, PublishFormat.@pfOneNote, "");
        }

        public void @FindPages(string @bstrStartNodeID, string @bstrSearchString, out string @pbstrHierarchyXmlOut, bool @fIncludeUnindexedPages, bool @fDisplay)
        {
            @FindPages(@bstrStartNodeID, @bstrSearchString, out @pbstrHierarchyXmlOut, @fIncludeUnindexedPages, @fDisplay, XMLSchema.@xs2013);
        }

        public void @FindPages(string @bstrStartNodeID, string @bstrSearchString, out string @pbstrHierarchyXmlOut, bool @fIncludeUnindexedPages)
        {
            @FindPages(@bstrStartNodeID, @bstrSearchString, out @pbstrHierarchyXmlOut, @fIncludeUnindexedPages, false, XMLSchema.@xs2013);
        }

        public void @FindPages(string @bstrStartNodeID, string @bstrSearchString, out string @pbstrHierarchyXmlOut)
        {
            @FindPages(@bstrStartNodeID, @bstrSearchString, out @pbstrHierarchyXmlOut, false, false, XMLSchema.@xs2013);
        }

        public void @FindMeta(string @bstrStartNodeID, string @bstrSearchStringName, out string @pbstrHierarchyXmlOut, bool @fIncludeUnindexedPages)
        {
            @FindMeta(@bstrStartNodeID, @bstrSearchStringName, out @pbstrHierarchyXmlOut, @fIncludeUnindexedPages, XMLSchema.@xs2013);
        }

        public void @FindMeta(string @bstrStartNodeID, string @bstrSearchStringName, out string @pbstrHierarchyXmlOut)
        {
            @FindMeta(@bstrStartNodeID, @bstrSearchStringName, out @pbstrHierarchyXmlOut, false, XMLSchema.@xs2013);
        }

    }
}

