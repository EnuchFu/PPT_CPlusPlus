// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

//#import "D:\\Office2013\\Office15\\MSPPT.OLB" no_namespace
// CTextRange 包装器类

class CTextRange : public COleDispatchDriver
{
public:
	CTextRange(){} // 调用 COleDispatchDriver 默认构造函数
	CTextRange(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CTextRange(const CTextRange& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// TextRange 方法
public:
	LPUNKNOWN get__NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}
	VARIANT _Index(long Index)
	{
		VARIANT result;
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0xa, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, Index);
		return result;
	}
	long get_Count()
	{
		long result;
		InvokeHelper(0xb, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Application()
	{
		LPDISPATCH result;
		InvokeHelper(0x7d1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Parent()
	{
		LPDISPATCH result;
		InvokeHelper(0x7d2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ActionSettings()
	{
		LPDISPATCH result;
		InvokeHelper(0x7d3, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Start()
	{
		long result;
		InvokeHelper(0x7d4, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long get_Length()
	{
		long result;
		InvokeHelper(0x7d5, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	float get_BoundLeft()
	{
		float result;
		InvokeHelper(0x7d6, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	float get_BoundTop()
	{
		float result;
		InvokeHelper(0x7d7, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	float get_BoundWidth()
	{
		float result;
		InvokeHelper(0x7d8, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	float get_BoundHeight()
	{
		float result;
		InvokeHelper(0x7d9, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH Paragraphs(long Start, long Length)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 VTS_I4;
		InvokeHelper(0x7da, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Start, Length);
		return result;
	}
	LPDISPATCH Sentences(long Start, long Length)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 VTS_I4;
		InvokeHelper(0x7db, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Start, Length);
		return result;
	}
	LPDISPATCH Words(long Start, long Length)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 VTS_I4;
		InvokeHelper(0x7dc, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Start, Length);
		return result;
	}
	LPDISPATCH Characters(long Start, long Length)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 VTS_I4;
		InvokeHelper(0x7dd, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Start, Length);
		return result;
	}
	LPDISPATCH Lines(long Start, long Length)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 VTS_I4;
		InvokeHelper(0x7de, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Start, Length);
		return result;
	}
	LPDISPATCH Runs(long Start, long Length)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 VTS_I4;
		InvokeHelper(0x7df, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Start, Length);
		return result;
	}
	LPDISPATCH TrimText()
	{
		LPDISPATCH result;
		InvokeHelper(0x7e0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	CString get_Text()
	{
		CString result;
		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_Text(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x0, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH InsertAfter(LPCTSTR NewText)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x7e1, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, NewText);
		return result;
	}
	LPDISPATCH InsertBefore(LPCTSTR NewText)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x7e2, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, NewText);
		return result;
	}
	LPDISPATCH InsertDateTime(long DateTimeFormat, long InsertAsField)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 VTS_I4;
		InvokeHelper(0x7e3, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, DateTimeFormat, InsertAsField);
		return result;
	}
	LPDISPATCH InsertSlideNumber()
	{
		LPDISPATCH result;
		InvokeHelper(0x7e4, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH InsertSymbol(LPCTSTR FontName, long CharNumber, long Unicode)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_I4 VTS_I4;
		InvokeHelper(0x7e5, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, FontName, CharNumber, Unicode);
		return result;
	}
	LPDISPATCH get_Font()
	{
		LPDISPATCH result;
		InvokeHelper(0x7e6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ParagraphFormat()
	{
		LPDISPATCH result;
		InvokeHelper(0x7e7, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_IndentLevel()
	{
		long result;
		InvokeHelper(0x7e8, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_IndentLevel(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x7e8, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	void Select()
	{
		InvokeHelper(0x7e9, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Cut()
	{
		InvokeHelper(0x7ea, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Copy()
	{
		InvokeHelper(0x7eb, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Delete()
	{
		InvokeHelper(0x7ec, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH Paste()
	{
		LPDISPATCH result;
		InvokeHelper(0x7ed, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void ChangeCase(long Type)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x7ee, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Type);
	}
	void AddPeriods()
	{
		InvokeHelper(0x7ef, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void RemovePeriods()
	{
		InvokeHelper(0x7f0, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH Find(LPCTSTR FindWhat, long After, long MatchCase, long WholeWords)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_I4 VTS_I4 VTS_I4;
		InvokeHelper(0x7f1, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, FindWhat, After, MatchCase, WholeWords);
		return result;
	}
	LPDISPATCH Replace(LPCTSTR FindWhat, LPCTSTR ReplaceWhat, long After, long MatchCase, long WholeWords)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_I4 VTS_I4 VTS_I4;
		InvokeHelper(0x7f2, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, FindWhat, ReplaceWhat, After, MatchCase, WholeWords);
		return result;
	}
	void RotatedBounds(float * X1, float * Y1, float * X2, float * Y2, float * X3, float * Y3, float * x4, float * y4)
	{
		static BYTE parms[] = VTS_PR4 VTS_PR4 VTS_PR4 VTS_PR4 VTS_PR4 VTS_PR4 VTS_PR4 VTS_PR4;
		InvokeHelper(0x7f3, DISPATCH_METHOD, VT_EMPTY, NULL, parms, X1, Y1, X2, Y2, X3, Y3, x4, y4);
	}
	long get_LanguageID()
	{
		long result;
		InvokeHelper(0x7f4, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_LanguageID(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x7f4, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	void RtlRun()
	{
		InvokeHelper(0x7f5, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void LtrRun()
	{
		InvokeHelper(0x7f6, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH PasteSpecial(long DataType, long DisplayAsIcon, LPCTSTR IconFileName, long IconIndex, LPCTSTR IconLabel, long Link)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 VTS_I4 VTS_BSTR VTS_I4 VTS_BSTR VTS_I4;
		InvokeHelper(0x7f7, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, DataType, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Link);
		return result;
	}

	// TextRange 属性
public:

};
