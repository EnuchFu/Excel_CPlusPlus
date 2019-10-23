#include "stdafx.h"
#include "My_Excel.h"
#include <io.h>

#define xlFormatFromLeftOrAbove 0 //从上方和/或左侧单元格复制格式
#define xlFormatFromRightOrBelow 1 //从下方和/或右侧单元格复制格式

CMyExcel::CMyExcel()
{
	m_UseRows = -1;
	m_UseCols = -1;
	m_StartRow = -1;
	m_StartCol = -1;

	m_ExcelApp = NULL;
	m_Excelbook = NULL;
	m_Excelrange = NULL;
	m_Excelbooks = NULL;
	m_Excelsheet = NULL;
	m_Excelsheets = NULL;
	lpDisp = NULL;

	m_filename = "";
}

CMyExcel::~CMyExcel()
{
	if (m_Excelsheets != NULL)
	{
		m_Excelsheets.ReleaseDispatch();
		m_Excelsheets = NULL;
	}

	if (m_Excelbooks != NULL)
	{
		m_Excelbooks.Close();
		m_Excelbooks.ReleaseDispatch();
		m_Excelbooks = NULL;
	}
	if (m_ExcelApp != NULL)
	{
		m_ExcelApp.Quit();
		m_ExcelApp.ReleaseDispatch();
		m_ExcelApp = NULL;
	}
}

//初始化Excel
bool CMyExcel::InitExcel(bool isVisible)
{
	if (!m_ExcelApp.CreateDispatch(_T("Excel.Application"), NULL))
	{
		AfxMessageBox(_T("启动Excel服务器失败!"));
		return FALSE;
	}
	else
	{
		m_ExcelApp.put_Visible(isVisible);
		return TRUE;
	}
}

//判断是否在vector中
bool CMyExcel::IsIntoVector(string inString, std::vector<string> inStringVector)
{
	vector<string>::iterator result = find(inStringVector.begin(), inStringVector.end(), inString);
	if (result == inStringVector.end()) //没找到
		return false;
	else //找到
		return true;
}

//创建新的Excel
bool CMyExcel::CreateExcel(bool isVisible, string inSheetName)
{
	if (InitExcel(isVisible))
	{
		m_Excelbooks.AttachDispatch(m_ExcelApp.get_Workbooks());
		lpDisp = m_Excelbooks.Add(vtMissing);
		m_Excelbook.AttachDispatch(lpDisp);
		m_Excelsheet = m_Excelbook.get_ActiveSheet();//获取当前工作表
		m_Excelsheets.AttachDispatch(m_Excelbook.get_Worksheets());
		GetExcelSheetNames();
		if (strcmp(inSheetName.c_str(), "Sheet1") != 0)
		{
			m_Excelsheet.put_Name(inSheetName.c_str());
		}
		return true;
	}
	return false;
}

//打开指定Excel
bool CMyExcel::OpenExcel(string inExcelName, bool isVisible)
{
	if (InitExcel(isVisible))
	{
		//判断文件是否存在
		if (_access(inExcelName.c_str(), 0) == -1)
		{
			m_ExcelApp.Quit();
			m_ExcelApp = NULL;
			MessageBoxA(NULL, "指定打开的文件不存在!", "提示", MB_ICONEXCLAMATION);
			return FALSE;
		}

		//判断文件是否有写入权限
		if (_access(inExcelName.c_str(), 2) == -1)
		{
			m_ExcelApp.Quit();
			m_ExcelApp = NULL;
			MessageBoxA(NULL, "指定打开的文件没有写入权限!", "提示", MB_ICONEXCLAMATION);
			return false;
		}

		m_Excelbooks.AttachDispatch(m_ExcelApp.get_Workbooks());
		lpDisp = m_Excelbooks.Open(inExcelName.c_str(), vtMissing, vtMissing, vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing, vtMissing, vtMissing);
		m_Excelbook.AttachDispatch(lpDisp);
		m_Excelsheets.AttachDispatch(m_Excelbook.get_Worksheets());
		GetExcelSheetNames();
		this->m_filename = inExcelName;
		return true;
	}
	return false;
}

//另存Excel
bool CMyExcel::SaveAs(string inSavePath, bool isReplace)
{
	try
	{
		if (strcmp(inSavePath.c_str(), this->m_filename.c_str()) != 0)
		{
			m_ExcelApp.put_DisplayAlerts(!isReplace);
			m_Excelbook.SaveAs(COleVariant(inSavePath.c_str()), vtMissing,vtMissing, vtMissing,
				vtMissing, vtMissing, (long)0, vtMissing, vtMissing, vtMissing,vtMissing, vtMissing);
			this->m_filename = inSavePath;
			return true;
		}
		else
		{
			MessageBoxA(NULL, "输入路径和原路径相同!", "提示", MB_ICONEXCLAMATION);
			return false;
		}
	}
	catch (exception &e)
	{
		MessageBoxA(NULL, e.what(), "错误", MB_ICONEXCLAMATION);
		return false;
	}
	
}

//保存Excel
bool CMyExcel::SaveExcel()
{
	try
	{
		m_Excelbook.Save();
		return true;
	}
	catch (exception &e)
	{
		MessageBoxA(NULL, e.what(), "错误", MB_ICONEXCLAMATION);
		return false;
	}
}

//关闭Excel
void CMyExcel::CloseExcel()
{
	m_ExcelApp.put_DisplayAlerts(true);
	m_Excelsheets.ReleaseDispatch();
	m_Excelsheets = NULL;
	m_Excelbooks.Close();
	m_Excelbooks.ReleaseDispatch();
	m_Excelbooks = NULL;
	m_ExcelApp.Quit();
	m_ExcelApp.ReleaseDispatch();
	m_ExcelApp = NULL;
}

//获得Excel所有的Sheet名称
std::vector<string> CMyExcel::GetExcelSheetNames()
{
	std::vector<string> m_sheetnames;
	int sheet_num = m_Excelsheets.get_Count();
	for (int i = 1; i <= sheet_num; i++)
	{
		m_Excelsheet.AttachDispatch(m_Excelsheets.get_Item(_variant_t(i)));
		m_sheetnames.push_back(m_Excelsheet.get_Name().GetBuffer(0));
	}
	this->m_Sheetnames = m_sheetnames;
	return m_sheetnames;
}

//添加工作表
bool CMyExcel::AddWorkSheet(string inSheetname, int inIndex)
{
	if (!IsIntoVector(inSheetname, m_Sheetnames))
	{
		if (inIndex >= 1 && inIndex <= (int)m_Sheetnames.size())
		{
			m_Excelsheet = m_Excelsheets.Add(vtMissing, _variant_t(m_Excelsheets.get_Item(_variant_t(inIndex))), _variant_t((long)1), vtMissing);
		}
		else
		{
			m_Excelsheet = m_Excelsheets.Add(vtMissing, _variant_t(m_Excelsheets.get_Item(_variant_t(m_Sheetnames.size()))), _variant_t((long)1), vtMissing);
		}
		m_Excelsheet.put_Name(inSheetname.c_str());
		GetExcelSheetNames();
		return true;
	}
	else
	{
		MessageBoxA(NULL, "要添加的工作表名称已经存在!", "提示", MB_ICONEXCLAMATION);
	}
	return false;
}

//删除工作表
bool CMyExcel::DeleteWorkSheet(string inSheetname)
{
	if (IsIntoVector(inSheetname, m_Sheetnames) && m_Sheetnames.size() > 1)
	{
		m_Excelsheet.AttachDispatch(m_Excelsheets.get_Item(COleVariant(inSheetname.c_str())));
		m_Excelsheet.Delete();
		GetExcelSheetNames();
		return true;
	}
	else
	{
		MessageBoxA(NULL, "要删除的工作表名称不存在或表中只存在一个工作表!", "提示", MB_ICONEXCLAMATION);
	}
	return false;
}

//重命名工作表
bool CMyExcel::RenameExcelSheet(string inOldSheetName, string inNewSheetName)
{
	if (IsIntoVector(inOldSheetName, m_Sheetnames) && !IsIntoVector(inNewSheetName, m_Sheetnames))
	{
		m_Excelsheet.AttachDispatch(m_Excelsheets.get_Item(COleVariant(inOldSheetName.c_str())));
		m_Excelsheet.put_Name(inNewSheetName.c_str());
		GetExcelSheetNames();
		return true;
	}
	else
	{
		MessageBoxA(NULL, "要替换的工作表名称不存在或替换名已存在于Excel中!", "提示", MB_ICONEXCLAMATION);
	}
	return false;
}

//获取单元格内容
std::string CMyExcel::GetCellInfor(string inSheetName, int inRow, int inColumn)
{
	CString cell_infor = "";
	m_Excelsheet.AttachDispatch(m_Excelsheets.get_Item(COleVariant(inSheetName.c_str())));

	inRow = (0 == inRow) ? 1 : inRow;
	inColumn = (0 == inColumn) ? 1 : inColumn;
	inColumn = (inColumn > 26) ? (inColumn % 26) : (inColumn);
	cell_infor.Format("%c%d", inColumn + 'A' - 1, inRow);
	m_Excelrange = m_Excelsheet.get_Range(COleVariant(cell_infor), COleVariant(cell_infor));

	//获得单元格的内容
	COleVariant vResult = m_Excelrange.get_Value2();
	m_Excelrange.ReleaseDispatch();
	if (vResult.vt == VT_BSTR)
	{
		cell_infor = vResult.bstrVal;
	}
	else if (vResult.vt == VT_R8)
	{
		cell_infor.Format("%f", vResult.dblVal);

	}
	else if (vResult.vt == VT_DATE)
	{
		SYSTEMTIME	st;
		VariantTimeToSystemTime(vResult.date, &st);
		TCHAR time[256] = "";
		sprintf_s(time, _T("%02d-%02d-%02d %02d:%02d:%02d"), st.wYear, st.wMonth,
			st.wDay, st.wHour, st.wMinute, st.wSecond);
		cell_infor = time;
	}
	else if (vResult.vt == VT_EMPTY)
	{
		cell_infor = " ";
	}
	return cell_infor.GetString();
}
std::string CMyExcel::GetCellInfor(string inSheetName, int inRow, char inColumn)
{
	return GetCellInfor(inSheetName, inRow, (int)inColumn - 64);
}

//设置单元格内容
void CMyExcel::SetCellValue(string inSheetName, int inRow, int inColumn, string inValue)
{
	CString cell_infor = "";
	m_Excelsheet.AttachDispatch(m_Excelsheets.get_Item(COleVariant(inSheetName.c_str())));

	inRow = (0 == inRow) ? 1 : inRow;
	inColumn = (0 == inColumn) ? 1 : inColumn;
	inColumn = (inColumn > 26) ? (inColumn % 26) : (inColumn);
	cell_infor.Format("%c%d", inColumn + 'A' - 1, inRow);
	m_Excelrange = m_Excelsheet.get_Range(COleVariant(cell_infor), COleVariant(cell_infor));
	m_Excelrange.put_Value2(COleVariant(inValue.c_str()));
	m_Excelrange.ReleaseDispatch();
}
void CMyExcel::SetCellValue(string inSheetName, int inRow, char inColumn, string inValue)
{
	SetCellValue(inSheetName, inRow, (int)inColumn - 64, inValue);
}

//设置工作表指定行的行高
void CMyExcel::SetRowHeight(string inSheetName, int inRow, int inHeightValue /*= 20*/)
{
	m_Excelsheet.AttachDispatch(m_Excelsheets.get_Item(COleVariant(inSheetName.c_str())));

	CString cell_infor = "";
	inRow = (0 == inRow) ? 1 : inRow;
	cell_infor.Format("%c%d", 'A', inRow);
	m_Excelrange = m_Excelsheet.get_Range(COleVariant(cell_infor), COleVariant(cell_infor));
	m_Excelrange.put_RowHeight(COleVariant((long)inHeightValue));
	m_Excelrange.ReleaseDispatch();
}

//设置工作表指定列的列宽
void CMyExcel::SetColumnWidth(string inSheetName, int inColumn, int inWidthValue /*= 20*/)
{	
	m_Excelsheet.AttachDispatch(m_Excelsheets.get_Item(COleVariant(inSheetName.c_str())));

	CString cell_infor = "";
	inColumn = (0 == inColumn) ? 1 : inColumn;
	inColumn = (inColumn > 26) ? (inColumn % 26) : (inColumn);
	cell_infor.Format("%c%d", inColumn + 'A' - 1, 1);
	m_Excelrange = m_Excelsheet.get_Range(COleVariant(cell_infor), COleVariant(cell_infor));
	m_Excelrange.put_ColumnWidth(COleVariant((long)inWidthValue));
	m_Excelrange.ReleaseDispatch();
}
void CMyExcel::SetColumnWidth(string inSheetName, char inColumn, int inWidthValue /*= 20*/)
{
	SetColumnWidth(inSheetName, (int)inColumn - 64, inWidthValue);
}

//插入一行
void CMyExcel::InsertRow(string inSheetName, int inRow)
{
	m_Excelsheet.AttachDispatch(m_Excelsheets.get_Item(COleVariant(inSheetName.c_str())));

	CString cell_infor = "";
	inRow = (0 == inRow) ? 1 : inRow;
	cell_infor.Format("%c%d", 'A', inRow);
	m_Excelrange = m_Excelsheet.get_Range(COleVariant(cell_infor), COleVariant(cell_infor));
	m_Excelrange.AttachDispatch(m_Excelrange.get_EntireRow());
	m_Excelrange.Insert(vtMissing, _variant_t(xlFormatFromLeftOrAbove));
	m_Excelrange.ReleaseDispatch();
}

//插入一列
void CMyExcel::InsertColumn(string inSheetName, int inColumn)
{
	m_Excelsheet.AttachDispatch(m_Excelsheets.get_Item(COleVariant(inSheetName.c_str())));

	CString cell_infor = "";
	inColumn = (0 == inColumn) ? 1 : inColumn;
	inColumn = (inColumn > 26) ? (inColumn % 26) : (inColumn);
	cell_infor.Format("%c%d", inColumn + 'A' - 1, 1);
	m_Excelrange = m_Excelsheet.get_Range(COleVariant(cell_infor), COleVariant(cell_infor));
	m_Excelrange.AttachDispatch(m_Excelrange.get_EntireColumn());
	m_Excelrange.Insert(vtMissing, _variant_t(xlFormatFromLeftOrAbove));
	m_Excelrange.ReleaseDispatch();
}
void CMyExcel::InsertColumn(string inSheetName, char inColumn)
{
	InsertColumn(inSheetName, (int)inColumn - 64);
}

//合并单元格
void CMyExcel::CombineRanges(string inSheetName, string left_top, string right_low)
{
	m_Excelsheet.AttachDispatch(m_Excelsheets.get_Item(COleVariant(inSheetName.c_str())));
	lpDisp = m_Excelsheet.get_Range(_variant_t(left_top.c_str()), _variant_t(right_low.c_str()));
	m_Excelrange.AttachDispatch(lpDisp, TRUE);
	m_Excelrange.Merge(_variant_t(0));//合并单元格
}
void CMyExcel::CombineRanges(string inSheetName, string left_top, string right_low, string inValue)
{
	m_Excelsheet.AttachDispatch(m_Excelsheets.get_Item(COleVariant(inSheetName.c_str())));
	lpDisp = m_Excelsheet.get_Range(_variant_t(left_top.c_str()), _variant_t(right_low.c_str()));
	m_Excelrange.AttachDispatch(lpDisp, TRUE);
	m_Excelrange.Merge(_variant_t(0));//合并单元格
	m_Excelrange.put_Value2(COleVariant(inValue.c_str()));
	m_Excelrange.ReleaseDispatch();
}

//设置单元格的颜色
void CMyExcel::SetCellColor(string inSheetName, int inRow, int inColumn, char RGBY)
{
	CString cell_infor = "";
	m_Excelsheet.AttachDispatch(m_Excelsheets.get_Item(COleVariant(inSheetName.c_str())));

	inRow = (0 == inRow) ? 1 : inRow;
	inColumn = (0 == inColumn) ? 1 : inColumn;
	inColumn = (inColumn > 26) ? (inColumn % 26) : (inColumn);
	cell_infor.Format("%c%d", inColumn + 'A' - 1, inRow);
	m_Excelrange = m_Excelsheet.get_Range(COleVariant(cell_infor), COleVariant(cell_infor));
	Cnterior interior = m_Excelrange.get_Interior();
	if (RGBY == 'R' || RGBY == 'r')
	{
		interior.put_Color(COleVariant((long)65536 * 0 + 256 * 0 + 255));
	}
	else if (RGBY == 'G' || RGBY == 'g')
	{
		interior.put_Color(COleVariant((long)65536 * 0 + 256 * 255 + 0));
	}
	else if (RGBY == 'B' || RGBY == 'b')
	{
		interior.put_Color(COleVariant((long)65536 * 255 + 256 * 0 + 0));
	}
	else if (RGBY == 'Y' || RGBY == 'y')
	{
		interior.put_Color(COleVariant((long)65536 * 0 + 256 * 255 + 255));
	}
	m_Excelrange.ReleaseDispatch();
}
void CMyExcel::SetCellColor(string inSheetName, int inRow, char inColumn, char RGBY)
{
	SetCellColor(inSheetName, inRow, (int)inColumn - 64, RGBY);
}

//改变表格指定单元格文字颜色
void CMyExcel::ChangeCellTextColor(string inSheetName, int inRow, int inColumn, char RGBY)
{
	CString cell_infor = "";
	m_Excelsheet.AttachDispatch(m_Excelsheets.get_Item(COleVariant(inSheetName.c_str())));

	inRow = (0 == inRow) ? 1 : inRow;
	inColumn = (0 == inColumn) ? 1 : inColumn;
	inColumn = (inColumn > 26) ? (inColumn % 26) : (inColumn);
	cell_infor.Format("%c%d", inColumn + 'A' - 1, inRow);
	m_Excelrange = m_Excelsheet.get_Range(COleVariant(cell_infor), COleVariant(cell_infor));
	CFont0 range_font = m_Excelrange.get_Font();
	if (RGBY == 'R' || RGBY == 'r')
	{
		range_font.put_Color(COleVariant((long)65536 * 0 + 256 * 0 + 255));
	}
	else if (RGBY == 'G' || RGBY == 'g')
	{
		range_font.put_Color(COleVariant((long)65536 * 0 + 256 * 255 + 0));
	}
	else if (RGBY == 'B' || RGBY == 'b')
	{
		range_font.put_Color(COleVariant((long)65536 * 255 + 256 * 0 + 0));
	}
	else if (RGBY == 'Y' || RGBY == 'y')
	{
		range_font.put_Color(COleVariant((long)65536 * 0 + 256 * 255 + 255));
	}
	range_font.ReleaseDispatch();
	m_Excelrange.ReleaseDispatch();
}
void CMyExcel::ChangeCellTextColor(string inSheetName, int inRow, char inColumn, char inRGBY)
{
	ChangeCellTextColor(inSheetName, inRow, (int)inColumn - 64, inRGBY);
}

//插入图片到单元格
void CMyExcel::InsertPicture(string inSheetName, int inRow, int inColumn, string inPicturePath, bool isDelete/* = true*/)
{
	CString cell_infor = "";
	m_Excelsheet.AttachDispatch(m_Excelsheets.get_Item(COleVariant(inSheetName.c_str())));

	inRow = (0 == inRow) ? 1 : inRow;
	inColumn = (0 == inColumn) ? 1 : inColumn;
	inColumn = (inColumn > 26) ? (inColumn % 26) : (inColumn);
	cell_infor.Format("%c%d", inColumn + 'A' - 1, inRow);
	m_Excelrange = m_Excelsheet.get_Range(COleVariant(cell_infor), COleVariant(cell_infor));
	float fleft = _variant_t(m_Excelrange.get_Left());
	float ftop = _variant_t(m_Excelrange.get_Top());
	float fwhidth = _variant_t(m_Excelrange.get_Width());
	float fheight = _variant_t(m_Excelrange.get_Height());

	CShapes shapes = m_Excelsheet.get_Shapes();
	if (_access(inPicturePath.c_str(), 0) != -1)
	{
		shapes.AddPicture(inPicturePath.c_str(), TRUE, TRUE, fleft, ftop, fwhidth, fheight);
		if (isDelete)
		{
			::DeleteFile(inPicturePath.c_str());
		}
	}
	else
	{
		MessageBoxA(NULL, "未找到需要插入的图片!", "提示", MB_ICONEXCLAMATION);
	}
	shapes.ReleaseDispatch();
	m_Excelrange.ReleaseDispatch();
}
void CMyExcel::InsertPicture(string inSheetName, int inRow, char inColumn, string inPicturePath, bool isDelete /*= true*/)
{
	InsertPicture(inSheetName, inRow, (int)inColumn - 64, inPicturePath, isDelete);
}
void CMyExcel::InsertPicture(string inSheetName, string left_top, string right_low, string inPicturePath, bool isDelete /*= true*/)
{
	m_Excelsheet.AttachDispatch(m_Excelsheets.get_Item(COleVariant(inSheetName.c_str())));
	lpDisp = m_Excelsheet.get_Range(_variant_t(left_top.c_str()), _variant_t(right_low.c_str()));
	m_Excelrange.AttachDispatch(lpDisp, TRUE);

	float fleft = _variant_t(m_Excelrange.get_Left());
	float ftop = _variant_t(m_Excelrange.get_Top());
	float fwhidth = _variant_t(m_Excelrange.get_Width());
	float fheight = _variant_t(m_Excelrange.get_Height());

	CShapes shapes = m_Excelsheet.get_Shapes();
	if (_access(inPicturePath.c_str(), 0) != -1)
	{
		shapes.AddPicture(inPicturePath.c_str(), TRUE, TRUE, fleft, ftop, fwhidth, fheight);
		if (isDelete)
		{
			::DeleteFile(inPicturePath.c_str());
		}
	}
	else
	{
		MessageBoxA(NULL, "未找到需要插入的图片!", "提示", MB_ICONEXCLAMATION);
	}
	shapes.ReleaseDispatch();
	m_Excelrange.ReleaseDispatch();

}

//设置单元格为文本类型
void CMyExcel::SetRangeFormat(string inSheetName, int inRow, int inColumn)
{
	CString cell_infor = "";
	m_Excelsheet.AttachDispatch(m_Excelsheets.get_Item(COleVariant(inSheetName.c_str())));

	inRow = (0 == inRow) ? 1 : inRow;
	inColumn = (0 == inColumn) ? 1 : inColumn;
	inColumn = (inColumn > 26) ? (inColumn % 26) : (inColumn);
	cell_infor.Format("%c%d", inColumn + 'A' - 1, inRow);
	m_Excelrange = m_Excelsheet.get_Range(COleVariant(cell_infor), COleVariant(cell_infor));
	m_Excelrange.put_NumberFormatLocal((_variant_t)"@");
	m_Excelrange.ReleaseDispatch();
}
void CMyExcel::SetRangeFormat(string inSheetName, int inRow, char inColumn)
{
	SetRangeFormat(inSheetName, inRow, (int)inColumn - 64);
}
void CMyExcel::SetRangeFormat(string inSheetName, string left_top, string right_low)
{
	m_Excelsheet.AttachDispatch(m_Excelsheets.get_Item(COleVariant(inSheetName.c_str())));
	lpDisp = m_Excelsheet.get_Range(_variant_t(left_top.c_str()), _variant_t(right_low.c_str()));
	m_Excelrange.AttachDispatch(lpDisp, TRUE);
	m_Excelrange.put_NumberFormatLocal((_variant_t)"@");
	m_Excelrange.ReleaseDispatch();
}

//设置单元格/区域 字体加粗
void CMyExcel::SetFontBold(string inSheetName, int inRow, int inColumn, bool isBold /*= true*/)
{
	CString cell_infor = "";
	m_Excelsheet.AttachDispatch(m_Excelsheets.get_Item(COleVariant(inSheetName.c_str())));

	inRow = (0 == inRow) ? 1 : inRow;
	inColumn = (0 == inColumn) ? 1 : inColumn;
	inColumn = (inColumn > 26) ? (inColumn % 26) : (inColumn);
	cell_infor.Format("%c%d", inColumn + 'A' - 1, inRow);
	m_Excelrange = m_Excelsheet.get_Range(COleVariant(cell_infor), COleVariant(cell_infor));

	CFont0 range_font = m_Excelrange.get_Font();
	if (isBold)
	{
		range_font.put_Bold(COleVariant((long)1));// 1：粗体，0：非粗体
	}
	else
	{
		range_font.put_Bold(COleVariant((long)0));// 1：粗体，0：非粗体
	}
	range_font.ReleaseDispatch();
	m_Excelrange.ReleaseDispatch();
}
void CMyExcel::SetFontBold(string inSheetName, int inRow, char inColumn, bool isBold /*= true*/)
{
	SetFontBold(inSheetName, inRow, (int)inColumn - 64, isBold);
}
void CMyExcel::SetFontBold(string inSheetName, string left_top, string right_low, bool isBold /*= true*/)
{
	m_Excelsheet.AttachDispatch(m_Excelsheets.get_Item(COleVariant(inSheetName.c_str())));
	lpDisp = m_Excelsheet.get_Range(_variant_t(left_top.c_str()), _variant_t(right_low.c_str()));
	m_Excelrange.AttachDispatch(lpDisp, TRUE);

	CFont0 range_font = m_Excelrange.get_Font();
	if (isBold)
	{
		range_font.put_Bold(COleVariant((long)1));// 1：粗体，0：非粗体
	}
	else
	{
		range_font.put_Bold(COleVariant((long)0));// 1：粗体，0：非粗体
	}
	range_font.ReleaseDispatch();
	m_Excelrange.ReleaseDispatch();
}

//获得工作表已使用的行列数
bool CMyExcel::GetWorkbookRC(string inSheetName)
{
	if (IsIntoVector(inSheetName, m_Sheetnames))
	{
		m_Excelsheet.AttachDispatch(m_Excelsheets.get_Item(COleVariant(inSheetName.c_str())));

		//取得用户区
		CRange userRange, getRange;
		userRange.AttachDispatch(m_Excelsheet.get_UsedRange());

		//得到用户区的行数
		getRange.AttachDispatch(userRange.get_Rows());
		this->m_UseRows = getRange.get_Count();

		//得到用户区的列数
		getRange.AttachDispatch(userRange.get_Columns());
		this->m_UseCols = getRange.get_Count();

		//得到用户区的开始行和开始列
		this->m_StartRow = userRange.get_Row();
		this->m_StartCol = userRange.get_Column();

		this->m_UseRows += this->m_StartRow - 1;
		this->m_UseCols += this->m_StartCol - 1;
		userRange.ReleaseDispatch();
		getRange.ReleaseDispatch();
		return true;
	}
	else
	{
		MessageBoxA(NULL, "指定名称的工作表不存在!", "提示", MB_ICONEXCLAMATION);
	}
	return false;
}

