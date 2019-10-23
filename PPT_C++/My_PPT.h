#pragma once
#include "CApplication.h"
#include "CSlides.h"
#include "CSlide.h"
#include "CPresentation.h"
#include "CPresentations.h"
#include "CSlideShowSettings.h"
#include "CShape.h"
#include "CShapes.h"
#include "CTextFrame.h"
#include "CTextRange.h"
#include "CTable0.h"
#include "CCell.h"
#include "CFillFormat.h"
#include "CColorFormat.h"
#include "CFonts.h"
#include "CFont0.h"
#include "CColumn.h"
#include "CColumns0.h"
#include "CRow.h"
#include "CRows.h"
#include <string>
#include <vector>

using namespace std;

class CMyPPT
{
public:
	CMyPPT();
	~CMyPPT();

	//新建PPT
	bool CreateNewPPT();

	//新建模板PPT
	bool CreatePPTofTemplate(string inTemplateFileName);

	//打开PPT
	bool OpenPPT(string inFileName, bool isVisible = true, bool isOnlyReadbool = false);

	//播放PPT	是否全屏：isFullScreen
	bool PlayPPT(bool isFullScreen);
	
	//添加页PPT	位置：inIndex	版式：inLayout
	bool AddSlide(int inIndex, int inLayout = 1);

	//删除页PPT
	bool DeleteSlide(int inIndex);

	//插入文本框
	CTextFrame InsertTextBox(int inIndex, string inTextString, float inLeft, float inTop, float inWidth = 1000.0, float inHeight = 1000.0);

	//改变文本框的内容
	bool ChangeTextBox(int inIndex, string inBoxName, string inTextString);

	//读取整个表格内容
	bool ReadTablePPT(int inIndex, string inTableName, int inRow, int inColumn, std::vector<std::vector<std::string>> &outTableInfor);

	//读取表格指定单元内容
	string ReadTablePPT(int inIndex, string inTableName, int inRow, int inColumn);

	//插入表格PPT
	CTable0 InsertTablePPT(int inIndex, int inRow, int inColumn, float inLeft, float inTop, float inWidth = 150.0, float inHeight = 20.0);
	
	//改变表格指定单元背景
	bool ChangeCellColor(int inIndex, string inTableName, int inRow, int inColumn, string inRGBY);

	//改变表格指定单元格文字颜色
	bool ChangeCellTextColor(int inIndex, string inTableName, int inRow, int inColumn, string inRGBY);

	//数据写入表格单元
	bool WriteStringToTable(int inIndex, string inTableName, int inRow, int inColumn, string inString);

	//表格添加一行
	bool AddRowToTable(int inIndex, string inTableName, int inRow = -1);

	//表格添加一列
	bool AddColumnToTable(int inIndex, string inTableName, int inColumn = -1);

	//插入图片PPT
	CShape InsertPicture(int inIndex, string inPicturePath, float inLeft, float inTop, float inWidth = 400.0, float inHeight = 300.0);

	//替换图片PPT
	CShape ReplacePicture(int inIndex, string inOldPicturePath, string inNewPicturePath);
	CShape ReplacePicture(int inIndex, CShape inShape, string inNewPicturePath);

	//删除指定的Shape
	bool DeleteShape(int inIndex, string inShapeName);

	//保存PPT
	bool SavePPT();

	//另存PPT
	bool SaveAsPPT(string inFullName);

	//关闭PPT
	bool ClosePPT();

	//关闭PPT(直接关闭，不保存)
	bool KillPPT();

private:
	CApplication		m_PPTApp;
	CSlides				m_Slides;
	CSlide				m_curSlide;
	CPresentation		m_Presentaion;
	CPresentations		m_Presentaions;
	CSlideShowSettings	m_slideshow;

	string m_filename;
	bool m_isvisible;
	int m_slidesnum;

	//初始化
	bool InitPPT();

};

