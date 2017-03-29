// Machine generated IDispatch wrapper class(es) created with ClassWizard
/////////////////////////////////////////////////////////////////////////////
// _Worksheet wrapper class

class _Worksheet : public COleDispatchDriver
{
public:
	_Worksheet() {}		// Calls COleDispatchDriver default constructor
	_Worksheet(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	_Worksheet(const _Worksheet& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	void Activate();
	void Copy(const VARIANT& Before, const VARIANT& After);
	void Delete();
	CString GetCodeName();
	CString Get_CodeName();
	void Set_CodeName(LPCTSTR lpszNewValue);
	long GetIndex();
	void Move(const VARIANT& Before, const VARIANT& After);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	LPDISPATCH GetNext();
	LPDISPATCH GetPageSetup();
	LPDISPATCH GetPrevious();
	void PrintPreview(const VARIANT& EnableChanges);
	BOOL GetProtectContents();
	BOOL GetProtectDrawingObjects();
	BOOL GetProtectionMode();
	BOOL GetProtectScenarios();
	void Select(const VARIANT& Replace);
	void Unprotect(const VARIANT& Password);
	long GetVisible();
	void SetVisible(long nNewValue);
	LPDISPATCH GetShapes();
	BOOL GetTransitionExpEval();
	void SetTransitionExpEval(BOOL bNewValue);
	BOOL GetAutoFilterMode();
	void SetAutoFilterMode(BOOL bNewValue);
	void SetBackgroundPicture(LPCTSTR Filename);
	void Calculate();
	BOOL GetEnableCalculation();
	void SetEnableCalculation(BOOL bNewValue);
	LPDISPATCH GetCells();
	LPDISPATCH ChartObjects(const VARIANT& Index);
	void CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& SpellLang);
	LPDISPATCH GetCircularReference();
	void ClearArrows();
	LPDISPATCH GetColumns();
	long GetConsolidationFunction();
	VARIANT GetConsolidationOptions();
	VARIANT GetConsolidationSources();
	BOOL GetEnableAutoFilter();
	void SetEnableAutoFilter(BOOL bNewValue);
	long GetEnableSelection();
	void SetEnableSelection(long nNewValue);
	BOOL GetEnableOutlining();
	void SetEnableOutlining(BOOL bNewValue);
	BOOL GetEnablePivotTable();
	void SetEnablePivotTable(BOOL bNewValue);
	VARIANT Evaluate(const VARIANT& Name);
	VARIANT _Evaluate(const VARIANT& Name);
	BOOL GetFilterMode();
	void ResetAllPageBreaks();
	LPDISPATCH GetNames();
	LPDISPATCH OLEObjects(const VARIANT& Index);
	LPDISPATCH GetOutline();
	void Paste(const VARIANT& Destination, const VARIANT& Link);
	LPDISPATCH PivotTables(const VARIANT& Index);
	LPDISPATCH PivotTableWizard(const VARIANT& SourceType, const VARIANT& SourceData, const VARIANT& TableDestination, const VARIANT& TableName, const VARIANT& RowGrand, const VARIANT& ColumnGrand, const VARIANT& SaveData, 
		const VARIANT& HasAutoFormat, const VARIANT& AutoPage, const VARIANT& Reserved, const VARIANT& BackgroundQuery, const VARIANT& OptimizeCache, const VARIANT& PageFieldOrder, const VARIANT& PageFieldWrapCount, const VARIANT& ReadData, 
		const VARIANT& Connection);
	LPDISPATCH GetRange(const VARIANT& Cell1, const VARIANT& Cell2);
	LPDISPATCH GetRows();
	LPDISPATCH Scenarios(const VARIANT& Index);
	CString GetScrollArea();
	void SetScrollArea(LPCTSTR lpszNewValue);
	void ShowAllData();
	void ShowDataForm();
	double GetStandardHeight();
	double GetStandardWidth();
	void SetStandardWidth(double newValue);
	BOOL GetTransitionFormEntry();
	void SetTransitionFormEntry(BOOL bNewValue);
	long GetType();
	LPDISPATCH GetUsedRange();
	LPDISPATCH GetHPageBreaks();
	LPDISPATCH GetVPageBreaks();
	LPDISPATCH GetQueryTables();
	BOOL GetDisplayPageBreaks();
	void SetDisplayPageBreaks(BOOL bNewValue);
	LPDISPATCH GetComments();
	LPDISPATCH GetHyperlinks();
	void ClearCircles();
	void CircleInvalid();
	LPDISPATCH GetAutoFilter();
	BOOL GetDisplayRightToLeft();
	void SetDisplayRightToLeft(BOOL bNewValue);
	LPDISPATCH GetScripts();
	void PrintOut(const VARIANT& From, const VARIANT& To, const VARIANT& Copies, const VARIANT& Preview, const VARIANT& ActivePrinter, const VARIANT& PrintToFile, const VARIANT& Collate, const VARIANT& PrToFileName);
	LPDISPATCH GetTab();
	LPDISPATCH GetMailEnvelope();
	void SaveAs(LPCTSTR Filename, const VARIANT& FileFormat, const VARIANT& Password, const VARIANT& WriteResPassword, const VARIANT& ReadOnlyRecommended, const VARIANT& CreateBackup, const VARIANT& AddToMru, const VARIANT& TextCodepage, 
		const VARIANT& TextVisualLayout, const VARIANT& Local);
	LPDISPATCH GetCustomProperties();
	LPDISPATCH GetSmartTags();
	LPDISPATCH GetProtection();
	void PasteSpecial(const VARIANT& Format, const VARIANT& Link, const VARIANT& DisplayAsIcon, const VARIANT& IconFileName, const VARIANT& IconIndex, const VARIANT& IconLabel, const VARIANT& NoHTMLFormatting);
	void Protect(const VARIANT& Password, const VARIANT& DrawingObjects, const VARIANT& Contents, const VARIANT& Scenarios, const VARIANT& UserInterfaceOnly, const VARIANT& AllowFormattingCells, const VARIANT& AllowFormattingColumns, 
		const VARIANT& AllowFormattingRows, const VARIANT& AllowInsertingColumns, const VARIANT& AllowInsertingRows, const VARIANT& AllowInsertingHyperlinks, const VARIANT& AllowDeletingColumns, const VARIANT& AllowDeletingRows, 
		const VARIANT& AllowSorting, const VARIANT& AllowFiltering, const VARIANT& AllowUsingPivotTables);
	LPDISPATCH GetListObjects();
	LPDISPATCH XmlDataQuery(LPCTSTR XPath, const VARIANT& SelectionNamespaces, const VARIANT& Map);
	LPDISPATCH XmlMapQuery(LPCTSTR XPath, const VARIANT& SelectionNamespaces, const VARIANT& Map);
};
/////////////////////////////////////////////////////////////////////////////
// _Workbook wrapper class

class _Workbook : public COleDispatchDriver
{
public:
	_Workbook() {}		// Calls COleDispatchDriver default constructor
	_Workbook(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	_Workbook(const _Workbook& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	BOOL GetAcceptLabelsInFormulas();
	void SetAcceptLabelsInFormulas(BOOL bNewValue);
	void Activate();
	LPDISPATCH GetActiveChart();
	LPDISPATCH GetActiveSheet();
	long GetAutoUpdateFrequency();
	void SetAutoUpdateFrequency(long nNewValue);
	BOOL GetAutoUpdateSaveChanges();
	void SetAutoUpdateSaveChanges(BOOL bNewValue);
	long GetChangeHistoryDuration();
	void SetChangeHistoryDuration(long nNewValue);
	LPDISPATCH GetBuiltinDocumentProperties();
	void ChangeFileAccess(long Mode, const VARIANT& WritePassword, const VARIANT& Notify);
	void ChangeLink(LPCTSTR Name, LPCTSTR NewName, long Type);
	LPDISPATCH GetCharts();
	void Close(const VARIANT& SaveChanges, const VARIANT& Filename, const VARIANT& RouteWorkbook);
	CString GetCodeName();
	CString Get_CodeName();
	void Set_CodeName(LPCTSTR lpszNewValue);
	VARIANT GetColors(const VARIANT& Index);
	void SetColors(const VARIANT& Index, const VARIANT& newValue);
	LPDISPATCH GetCommandBars();
	long GetConflictResolution();
	void SetConflictResolution(long nNewValue);
	LPDISPATCH GetContainer();
	BOOL GetCreateBackup();
	LPDISPATCH GetCustomDocumentProperties();
	BOOL GetDate1904();
	void SetDate1904(BOOL bNewValue);
	void DeleteNumberFormat(LPCTSTR NumberFormat);
	long GetDisplayDrawingObjects();
	void SetDisplayDrawingObjects(long nNewValue);
	BOOL ExclusiveAccess();
	long GetFileFormat();
	void ForwardMailer();
	CString GetFullName();
	BOOL GetHasPassword();
	BOOL GetHasRoutingSlip();
	void SetHasRoutingSlip(BOOL bNewValue);
	BOOL GetIsAddin();
	void SetIsAddin(BOOL bNewValue);
	VARIANT LinkInfo(LPCTSTR Name, long LinkInfo, const VARIANT& Type, const VARIANT& EditionRef);
	VARIANT LinkSources(const VARIANT& Type);
	LPDISPATCH GetMailer();
	void MergeWorkbook(const VARIANT& Filename);
	BOOL GetMultiUserEditing();
	CString GetName();
	LPDISPATCH GetNames();
	LPDISPATCH NewWindow();
	void OpenLinks(LPCTSTR Name, const VARIANT& ReadOnly, const VARIANT& Type);
	CString GetPath();
	BOOL GetPersonalViewListSettings();
	void SetPersonalViewListSettings(BOOL bNewValue);
	BOOL GetPersonalViewPrintSettings();
	void SetPersonalViewPrintSettings(BOOL bNewValue);
	LPDISPATCH PivotCaches();
	void Post(const VARIANT& DestName);
	BOOL GetPrecisionAsDisplayed();
	void SetPrecisionAsDisplayed(BOOL bNewValue);
	void PrintPreview(const VARIANT& EnableChanges);
	void ProtectSharing(const VARIANT& Filename, const VARIANT& Password, const VARIANT& WriteResPassword, const VARIANT& ReadOnlyRecommended, const VARIANT& CreateBackup, const VARIANT& SharingPassword);
	BOOL GetProtectStructure();
	BOOL GetProtectWindows();
	BOOL GetReadOnly();
	void RefreshAll();
	void Reply();
	void ReplyAll();
	void RemoveUser(long Index);
	long GetRevisionNumber();
	void Route();
	BOOL GetRouted();
	LPDISPATCH GetRoutingSlip();
	void RunAutoMacros(long Which);
	void Save();
	void SaveCopyAs(const VARIANT& Filename);
	BOOL GetSaved();
	void SetSaved(BOOL bNewValue);
	BOOL GetSaveLinkValues();
	void SetSaveLinkValues(BOOL bNewValue);
	void SendMail(const VARIANT& Recipients, const VARIANT& Subject, const VARIANT& ReturnReceipt);
	void SendMailer(const VARIANT& FileFormat, long Priority);
	void SetLinkOnData(LPCTSTR Name, const VARIANT& Procedure);
	LPDISPATCH GetSheets();
	BOOL GetShowConflictHistory();
	void SetShowConflictHistory(BOOL bNewValue);
	LPDISPATCH GetStyles();
	void Unprotect(const VARIANT& Password);
	void UnprotectSharing(const VARIANT& SharingPassword);
	void UpdateFromFile();
	void UpdateLink(const VARIANT& Name, const VARIANT& Type);
	BOOL GetUpdateRemoteReferences();
	void SetUpdateRemoteReferences(BOOL bNewValue);
	VARIANT GetUserStatus();
	LPDISPATCH GetCustomViews();
	LPDISPATCH GetWindows();
	LPDISPATCH GetWorksheets();
	BOOL GetWriteReserved();
	CString GetWriteReservedBy();
	LPDISPATCH GetExcel4IntlMacroSheets();
	LPDISPATCH GetExcel4MacroSheets();
	BOOL GetTemplateRemoveExtData();
	void SetTemplateRemoveExtData(BOOL bNewValue);
	void HighlightChangesOptions(const VARIANT& When, const VARIANT& Who, const VARIANT& Where);
	BOOL GetHighlightChangesOnScreen();
	void SetHighlightChangesOnScreen(BOOL bNewValue);
	BOOL GetKeepChangeHistory();
	void SetKeepChangeHistory(BOOL bNewValue);
	BOOL GetListChangesOnNewSheet();
	void SetListChangesOnNewSheet(BOOL bNewValue);
	void PurgeChangeHistoryNow(long Days, const VARIANT& SharingPassword);
	void AcceptAllChanges(const VARIANT& When, const VARIANT& Who, const VARIANT& Where);
	void RejectAllChanges(const VARIANT& When, const VARIANT& Who, const VARIANT& Where);
	void ResetColors();
	LPDISPATCH GetVBProject();
	void FollowHyperlink(LPCTSTR Address, const VARIANT& SubAddress, const VARIANT& NewWindow, const VARIANT& AddHistory, const VARIANT& ExtraInfo, const VARIANT& Method, const VARIANT& HeaderInfo);
	void AddToFavorites();
	BOOL GetIsInplace();
	void PrintOut(const VARIANT& From, const VARIANT& To, const VARIANT& Copies, const VARIANT& Preview, const VARIANT& ActivePrinter, const VARIANT& PrintToFile, const VARIANT& Collate, const VARIANT& PrToFileName);
	void WebPagePreview();
	LPDISPATCH GetPublishObjects();
	LPDISPATCH GetWebOptions();
	void ReloadAs(long Encoding);
	LPDISPATCH GetHTMLProject();
	BOOL GetEnvelopeVisible();
	void SetEnvelopeVisible(BOOL bNewValue);
	long GetCalculationVersion();
	BOOL GetVBASigned();
	BOOL GetShowPivotTableFieldList();
	void SetShowPivotTableFieldList(BOOL bNewValue);
	long GetUpdateLinks();
	void SetUpdateLinks(long nNewValue);
	void BreakLink(LPCTSTR Name, long Type);
	void SaveAs(const VARIANT& Filename, const VARIANT& FileFormat, const VARIANT& Password, const VARIANT& WriteResPassword, const VARIANT& ReadOnlyRecommended, const VARIANT& CreateBackup, long AccessMode, const VARIANT& ConflictResolution, 
		const VARIANT& AddToMru, const VARIANT& TextCodepage, const VARIANT& TextVisualLayout, const VARIANT& Local);
	BOOL GetEnableAutoRecover();
	void SetEnableAutoRecover(BOOL bNewValue);
	BOOL GetRemovePersonalInformation();
	void SetRemovePersonalInformation(BOOL bNewValue);
	CString GetFullNameURLEncoded();
	void CheckIn(const VARIANT& SaveChanges, const VARIANT& Comments, const VARIANT& MakePublic);
	BOOL CanCheckIn();
	void SendForReview(const VARIANT& Recipients, const VARIANT& Subject, const VARIANT& ShowMessage, const VARIANT& IncludeAttachment);
	void ReplyWithChanges(const VARIANT& ShowMessage);
	void EndReview();
	CString GetPassword();
	void SetPassword(LPCTSTR lpszNewValue);
	CString GetWritePassword();
	void SetWritePassword(LPCTSTR lpszNewValue);
	CString GetPasswordEncryptionProvider();
	CString GetPasswordEncryptionAlgorithm();
	long GetPasswordEncryptionKeyLength();
	void SetPasswordEncryptionOptions(const VARIANT& PasswordEncryptionProvider, const VARIANT& PasswordEncryptionAlgorithm, const VARIANT& PasswordEncryptionKeyLength, const VARIANT& PasswordEncryptionFileProperties);
	BOOL GetPasswordEncryptionFileProperties();
	BOOL GetReadOnlyRecommended();
	void SetReadOnlyRecommended(BOOL bNewValue);
	void Protect(const VARIANT& Password, const VARIANT& Structure, const VARIANT& Windows);
	LPDISPATCH GetSmartTagOptions();
	void RecheckSmartTags();
	LPDISPATCH GetPermission();
	LPDISPATCH GetSharedWorkspace();
	LPDISPATCH GetSync();
	void SendFaxOverInternet(const VARIANT& Recipients, const VARIANT& Subject, const VARIANT& ShowMessage);
	LPDISPATCH GetXmlNamespaces();
	LPDISPATCH GetXmlMaps();
	long XmlImport(LPCTSTR Url, LPDISPATCH* ImportMap, const VARIANT& Overwrite, const VARIANT& Destination);
	LPDISPATCH GetSmartDocument();
	LPDISPATCH GetDocumentLibraryVersions();
	BOOL GetInactiveListBorderVisible();
	void SetInactiveListBorderVisible(BOOL bNewValue);
	BOOL GetDisplayInkComments();
	void SetDisplayInkComments(BOOL bNewValue);
	long XmlImportXml(LPCTSTR Data, LPDISPATCH* ImportMap, const VARIANT& Overwrite, const VARIANT& Destination);
	void SaveAsXMLData(LPCTSTR Filename, LPDISPATCH Map);
	void ToggleFormsDesign();
};
/////////////////////////////////////////////////////////////////////////////
// Workbooks wrapper class

class Workbooks : public COleDispatchDriver
{
public:
	Workbooks() {}		// Calls COleDispatchDriver default constructor
	Workbooks(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Workbooks(const Workbooks& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH Add(const VARIANT& Template);
	void Close();
	long GetCount();
	LPDISPATCH GetItem(const VARIANT& Index);
	LPUNKNOWN Get_NewEnum();
	LPDISPATCH Get_Default(const VARIANT& Index);
	LPDISPATCH Open(LPCTSTR Filename, const VARIANT& UpdateLinks, const VARIANT& ReadOnly, const VARIANT& Format, const VARIANT& Password, const VARIANT& WriteResPassword, const VARIANT& IgnoreReadOnlyRecommended, const VARIANT& Origin, 
		const VARIANT& Delimiter, const VARIANT& Editable, const VARIANT& Notify, const VARIANT& Converter, const VARIANT& AddToMru, const VARIANT& Local, const VARIANT& CorruptLoad);
	void OpenText(LPCTSTR Filename, const VARIANT& Origin, const VARIANT& StartRow, const VARIANT& DataType, long TextQualifier, const VARIANT& ConsecutiveDelimiter, const VARIANT& Tab, const VARIANT& Semicolon, const VARIANT& Comma, 
		const VARIANT& Space, const VARIANT& Other, const VARIANT& OtherChar, const VARIANT& FieldInfo, const VARIANT& TextVisualLayout, const VARIANT& DecimalSeparator, const VARIANT& ThousandsSeparator, const VARIANT& TrailingMinusNumbers, 
		const VARIANT& Local);
	LPDISPATCH OpenDatabase(LPCTSTR Filename, const VARIANT& CommandText, const VARIANT& CommandType, const VARIANT& BackgroundQuery, const VARIANT& ImportDataAs);
	void CheckOut(LPCTSTR Filename);
	BOOL CanCheckOut(LPCTSTR Filename);
	LPDISPATCH OpenXML(LPCTSTR Filename, const VARIANT& Stylesheets, const VARIANT& LoadOption);
};
/////////////////////////////////////////////////////////////////////////////
// Worksheets wrapper class

class Worksheets : public COleDispatchDriver
{
public:
	Worksheets() {}		// Calls COleDispatchDriver default constructor
	Worksheets(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Worksheets(const Worksheets& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH Add(const VARIANT& Before, const VARIANT& After, const VARIANT& Count, const VARIANT& Type);
	void Copy(const VARIANT& Before, const VARIANT& After);
	long GetCount();
	void Delete();
	void FillAcrossSheets(LPDISPATCH Range, long Type);
	LPDISPATCH GetItem(const VARIANT& Index);
	void Move(const VARIANT& Before, const VARIANT& After);
	LPUNKNOWN Get_NewEnum();
	void PrintPreview(const VARIANT& EnableChanges);
	void Select(const VARIANT& Replace);
	LPDISPATCH GetHPageBreaks();
	LPDISPATCH GetVPageBreaks();
	VARIANT GetVisible();
	void SetVisible(const VARIANT& newValue);
	LPDISPATCH Get_Default(const VARIANT& Index);
	void PrintOut(const VARIANT& From, const VARIANT& To, const VARIANT& Copies, const VARIANT& Preview, const VARIANT& ActivePrinter, const VARIANT& PrintToFile, const VARIANT& Collate, const VARIANT& PrToFileName);
};
