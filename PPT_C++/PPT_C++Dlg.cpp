
// PPT_C++Dlg.cpp : 实现文件
//

#include "stdafx.h"
#include "PPT_C++.h"
#include "PPT_C++Dlg.h"
#include "afxdialogex.h"
#include "My_PPT.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CPPT_CDlg 对话框



CPPT_CDlg::CPPT_CDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CPPT_CDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CPPT_CDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CPPT_CDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDOK, &CPPT_CDlg::OnBnClickedOk)
END_MESSAGE_MAP()


// CPPT_CDlg 消息处理程序

BOOL CPPT_CDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO:  在此添加额外的初始化代码

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CPPT_CDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CPPT_CDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CPPT_CDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CPPT_CDlg::OnBnClickedOk()
{
	// TODO:  在此添加控件通知处理程序代码
	CMyPPT *fwj = new CMyPPT();
	fwj->OpenPPT("H:\\PAT_项目号_工具名称_Help文档_版本号_8位日期-正式模板.pptx", false);
	fwj->InsertTextBox(1, "烟火里的尘埃\n我想吃火锅", 20.0, 20.0);
	fwj->ChangeTextBox(14, "TextBox 5", "更改的内容");
	fwj->AddSlide(1);
	vector<vector<string>> kbe;
	fwj->ReadTablePPT(17, "Table 2", 5, 3, kbe);
	fwj->InsertTablePPT(1, 5, 10, 50, 50, 400, 10);
	fwj->WriteStringToTable(17, "Table 2", 2, 3, "23我想吃火锅");
	fwj->WriteStringToTable(17, "Table 2", 3, 2, "32烟火里的尘埃");
	fwj->ChangeCellColor(17, "Table 2", 3, 2, "B");
	fwj->ChangeCellTextColor(17, "Table 2", 2, 3, "Y");
	CShape oldpicture = fwj->InsertPicture(5, "G:\\图片记忆\\2.27交大\\2016-02-27_1.jpg", 50, 50, 400, 300);
	fwj->ReplacePicture(5, oldpicture, "G:\\图片记忆\\2.27交大\\2016-02-27_8.jpg");
	fwj->AddRowToTable(17, "Table 2");
	fwj->AddColumnToTable(17, "Table 2");
	fwj->AddColumnToTable(17, "Table 2", 3);
	fwj->SaveAsPPT("C:\\Users\\hp\\Desktop\\fwj.pptx");
	fwj->ClosePPT();
	//CDialogEx::OnOK();
}
