// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

//#import "D:\\Office2013\\Office15\\MSPPT.OLB" no_namespace
// CApplication 包装器类

class CApplication : public COleDispatchDriver
{
public:
	CApplication(){} // 调用 COleDispatchDriver 默认构造函数
	CApplication(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CApplication(const CApplication& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// _Application 方法
public:
	LPDISPATCH get_Presentations()
	{
		LPDISPATCH result;
		InvokeHelper(0x7d1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Windows()
	{
		LPDISPATCH result;
		InvokeHelper(0x7d2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPUNKNOWN get_Dialogs()
	{
		LPUNKNOWN result;
		InvokeHelper(0x7d3, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ActiveWindow()
	{
		LPDISPATCH result;
		InvokeHelper(0x7d4, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ActivePresentation()
	{
		LPDISPATCH result;
		InvokeHelper(0x7d5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_SlideShowWindows()
	{
		LPDISPATCH result;
		InvokeHelper(0x7d6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_CommandBars()
	{
		LPDISPATCH result;
		InvokeHelper(0x7d7, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	CString get_Path()
	{
		CString result;
		InvokeHelper(0x7d8, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	CString get_Name()
	{
		CString result;
		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	CString get_Caption()
	{
		CString result;
		InvokeHelper(0x7d9, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_Caption(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x7d9, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Assistant()
	{
		LPDISPATCH result;
		InvokeHelper(0x7da, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_FileSearch()
	{
		LPDISPATCH result;
		InvokeHelper(0x7db, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_FileFind()
	{
		LPDISPATCH result;
		InvokeHelper(0x7dc, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	CString get_Build()
	{
		CString result;
		InvokeHelper(0x7dd, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	CString get_Version()
	{
		CString result;
		InvokeHelper(0x7de, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	CString get_OperatingSystem()
	{
		CString result;
		InvokeHelper(0x7df, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	CString get_ActivePrinter()
	{
		CString result;
		InvokeHelper(0x7e0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	long get_Creator()
	{
		long result;
		InvokeHelper(0x7e1, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_AddIns()
	{
		LPDISPATCH result;
		InvokeHelper(0x7e2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_VBE()
	{
		LPDISPATCH result;
		InvokeHelper(0x7e3, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void Help(LPCTSTR HelpFile, long ContextID)
	{
		static BYTE parms[] = VTS_BSTR VTS_I4;
		InvokeHelper(0x7e4, DISPATCH_METHOD, VT_EMPTY, NULL, parms, HelpFile, ContextID);
	}
	void Quit()
	{
		InvokeHelper(0x7e5, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	VARIANT Run(LPCTSTR MacroName, SAFEARRAY * * safeArrayOfParams)
	{
		VARIANT result;
		static BYTE parms[] = VTS_BSTR VTS_UNKNOWN;
		InvokeHelper(0x7e6, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, MacroName, safeArrayOfParams);
		return result;
	}
	LPUNKNOWN PPFileDialog(long Type)
	{
		LPUNKNOWN result;
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x7e7, DISPATCH_METHOD, VT_UNKNOWN, (void*)&result, parms, Type);
		return result;
	}
	void LaunchSpelling(LPDISPATCH pWindow)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0x7e8, DISPATCH_METHOD, VT_EMPTY, NULL, parms, pWindow);
	}
	float get_Left()
	{
		float result;
		InvokeHelper(0x7e9, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_Left(float newValue)
	{
		static BYTE parms[] = VTS_R4;
		InvokeHelper(0x7e9, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_Top()
	{
		float result;
		InvokeHelper(0x7ea, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_Top(float newValue)
	{
		static BYTE parms[] = VTS_R4;
		InvokeHelper(0x7ea, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_Width()
	{
		float result;
		InvokeHelper(0x7eb, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_Width(float newValue)
	{
		static BYTE parms[] = VTS_R4;
		InvokeHelper(0x7eb, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_Height()
	{
		float result;
		InvokeHelper(0x7ec, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_Height(float newValue)
	{
		static BYTE parms[] = VTS_R4;
		InvokeHelper(0x7ec, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_WindowState()
	{
		long result;
		InvokeHelper(0x7ed, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_WindowState(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x7ed, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_Visible()
	{
		long result;
		InvokeHelper(0x7ee, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Visible(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x7ee, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_HWND()
	{
		long result;
		InvokeHelper(0x7ef, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long get_Active()
	{
		long result;
		InvokeHelper(0x7f0, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void Activate()
	{
		InvokeHelper(0x7f1, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH get_AnswerWizard()
	{
		LPDISPATCH result;
		InvokeHelper(0x7f2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_COMAddIns()
	{
		LPDISPATCH result;
		InvokeHelper(0x7f3, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	CString get_ProductCode()
	{
		CString result;
		InvokeHelper(0x7f4, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_DefaultWebOptions()
	{
		LPDISPATCH result;
		InvokeHelper(0x7f5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_LanguageSettings()
	{
		LPDISPATCH result;
		InvokeHelper(0x7f6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_MsoDebugOptions()
	{
		LPDISPATCH result;
		InvokeHelper(0x7f7, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_ShowWindowsInTaskbar()
	{
		long result;
		InvokeHelper(0x7f8, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_ShowWindowsInTaskbar(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x7f8, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPUNKNOWN get_Marker()
	{
		LPUNKNOWN result;
		InvokeHelper(0x7f9, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}
	long get_FeatureInstall()
	{
		long result;
		InvokeHelper(0x7fa, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_FeatureInstall(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x7fa, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL GetOptionFlag(long Option, BOOL Persist)
	{
		BOOL result;
		static BYTE parms[] = VTS_I4 VTS_BOOL;
		InvokeHelper(0x7fb, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, Option, Persist);
		return result;
	}
	void SetOptionFlag(long Option, BOOL State, BOOL Persist)
	{
		static BYTE parms[] = VTS_I4 VTS_BOOL VTS_BOOL;
		InvokeHelper(0x7fc, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Option, State, Persist);
	}
	LPDISPATCH get_FileDialog(long Type)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x7fd, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms, Type);
		return result;
	}
	long get_DisplayGridLines()
	{
		long result;
		InvokeHelper(0x7fe, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_DisplayGridLines(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x7fe, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_AutomationSecurity()
	{
		long result;
		InvokeHelper(0x7ff, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_AutomationSecurity(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x7ff, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_NewPresentation()
	{
		LPDISPATCH result;
		InvokeHelper(0x800, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_DisplayAlerts()
	{
		long result;
		InvokeHelper(0x801, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_DisplayAlerts(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x801, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_ShowStartupDialog()
	{
		long result;
		InvokeHelper(0x802, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_ShowStartupDialog(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x802, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	void SetPerfMarker(long Marker)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x803, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Marker);
	}
	LPDISPATCH get_AutoCorrect()
	{
		LPDISPATCH result;
		InvokeHelper(0x804, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Options()
	{
		LPDISPATCH result;
		InvokeHelper(0x805, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void LaunchPublishSlidesDialog(LPCTSTR SlideLibraryUrl)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x806, DISPATCH_METHOD, VT_EMPTY, NULL, parms, SlideLibraryUrl);
	}
	void LaunchSendToPPTDialog(VARIANT * SlideUrls)
	{
		static BYTE parms[] = VTS_PVARIANT;
		InvokeHelper(0x807, DISPATCH_METHOD, VT_EMPTY, NULL, parms, SlideUrls);
	}
	BOOL get_DisplayDocumentInformationPanel()
	{
		BOOL result;
		InvokeHelper(0x808, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_DisplayDocumentInformationPanel(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0x808, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Assistance()
	{
		LPDISPATCH result;
		InvokeHelper(0x809, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_ActiveEncryptionSession()
	{
		long result;
		InvokeHelper(0x80a, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_FileConverters()
	{
		LPDISPATCH result;
		InvokeHelper(0x80b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_SmartArtLayouts()
	{
		LPDISPATCH result;
		InvokeHelper(0x80c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_SmartArtQuickStyles()
	{
		LPDISPATCH result;
		InvokeHelper(0x80d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_SmartArtColors()
	{
		LPDISPATCH result;
		InvokeHelper(0x80e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ProtectedViewWindows()
	{
		LPDISPATCH result;
		InvokeHelper(0x80f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ActiveProtectedViewWindow()
	{
		LPDISPATCH result;
		InvokeHelper(0x810, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL get_IsSandboxed()
	{
		BOOL result;
		InvokeHelper(0x811, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ResampleMediaTasks()
	{
		LPDISPATCH result;
		InvokeHelper(0x812, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void StartNewUndoEntry()
	{
		InvokeHelper(0x813, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	long get_FileValidation()
	{
		long result;
		InvokeHelper(0x814, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_FileValidation(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x814, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH OpenThemeFile(LPCTSTR themeFileName)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x815, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, themeFileName);
		return result;
	}
	BOOL get_ChartDataPointTrack()
	{
		BOOL result;
		InvokeHelper(0x816, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_ChartDataPointTrack(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0x816, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_DisplayGuides()
	{
		long result;
		InvokeHelper(0x817, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_DisplayGuides(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x817, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}

	// _Application 属性
public:

};
