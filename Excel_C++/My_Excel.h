
#ifndef My_Excel_H_INCLUDED
#define My_Excel_H_INCLUDED

#pragma once

#include "CApplication.h"
#include "CBorders.h"
#include "CColorFormat.h"
#include "CFont0.h"
#include "CRange.h"
#include "CRanges.h"
#include "CShape.h"
#include "CShapes.h"
#include "CWorkbook.h"
#include "CWorksheet.h"
#include "CWorkbooks.h"
#include "CWorksheets.h"
#include "Cnterior.h"

#include <string>
#include <vector>

#define DllExport __declspec(dllexport)

using namespace std;

class DllExport CMyExcel
{
public:
	CMyExcel();
	~CMyExcel();

	//新建Excel
	bool CreateExcel(bool isVisible = false, string inSheetName = "Sheet1");

	//打开Excel
	bool OpenExcel(string inExcelName, bool isVisible = false);

	//另存Excel
	bool SaveAs(string inSavePath, bool isReplace = true);

	//保存Excel
	bool SaveExcel();

	//关闭Excel
	void CloseExcel();

	//添加工作表 插入到inIndex之后
	bool AddWorkSheet(string inSheetname, int inIndex = 1);

	//删除工作表
	bool DeleteWorkSheet(string inSheetname);

	//重命名工作表
	bool RenameExcelSheet(string inOldSheetName, string inNewSheetName);

	//获取单元格内容
	string GetCellInfor(string inSheetName, int inRow, int inColumn);
	string GetCellInfor(string inSheetName, int inRow, char inColumn);

	//设置单元格的值
	void SetCellValue(string inSheetName, int inRow, int inColumn, string inValue);
	void SetCellValue(string inSheetName, int inRow, char inColumn, string inValue);

	//设置工作表指定行的行高
	void SetRowHeight(string inSheetName, int inRow, int inHeightValue = 20);

	//设置工作表指定列的行宽
	void SetColumnWidth(string inSheetName, int inColumn, int inWidthValue = 15);
	void SetColumnWidth(string inSheetName, char inColumn, int inWidthValue = 15);

	//插入一行
	void InsertRow(string inSheetName, int inRow);

	//插入一列
	void InsertColumn(string inSheetName, int inColumn);
	void InsertColumn(string inSheetName, char inColumn);

	//合并单元格
	void CombineRanges(string inSheetName, string left_top, string right_low);
	void CombineRanges(string inSheetName, string left_top, string right_low, string inValue);

	//设置单元格的底色
	void SetCellColor(string inSheetName, int inRow, int inColumn, char inRGBY);
	void SetCellColor(string inSheetName, int inRow, char inColumn, char inRGBY);

	//改变表格指定单元格文字颜色
	void ChangeCellTextColor(string inSheetName, int inRow, int inColumn, char inRGBY);
	void ChangeCellTextColor(string inSheetName, int inRow, char inColumn, char inRGBY);

	//插入图片到指定单元格/区域
	void InsertPicture(string inSheetName, int inRow, int inColumn, string inPicturePath, bool isDelete = true);
	void InsertPicture(string inSheetName, int inRow, char inColumn, string inPicturePath, bool isDelete = true);
	void InsertPicture(string inSheetName, string left_top, string right_low, string inPicturePath, bool isDelete = true);

	//设置单元格为文本类型
	void SetRangeFormat(string inSheetName, int inRow, int inColumn);
	void SetRangeFormat(string inSheetName, int inRow, char inColumn);
	void SetRangeFormat(string inSheetName, string left_top, string right_low);

	//设置单元格/区域 字体加粗
	void SetFontBold(string inSheetName, int inRow, int inColumn, bool isBold = true);
	void SetFontBold(string inSheetName, int inRow, char inColumn, bool isBold = true);
	void SetFontBold(string inSheetName, string left_top, string right_low, bool isBold = true);

private:
	// 1, 2, 3, 4, 5, 6, 7, 8
	CApplication	m_ExcelApp;
	CWorkbook		m_Excelbook;
	CWorkbooks		m_Excelbooks;
	CWorksheet		m_Excelsheet;
	CWorksheets		m_Excelsheets;
	CRange			m_Excelrange;
	
	string m_filename;
	LPDISPATCH lpDisp;//打开Excel路径
	std::vector<string> m_Sheetnames;//获得所有工作表名称

	//以下四个参数暂时不放出来，需要的时候可以临时设置为public·······GetWorkbookRC()
	int m_StartRow ;//用户开始行
	int m_StartCol;//用户开始列
	int m_UseRows;//已使用的行数, 最大行
	int m_UseCols;//已使用的列数, 最大列

	//初始化
	bool InitExcel(bool isVisible);

	//判断是否在vector中
	bool IsIntoVector( string inString, std::vector<string> inStringVector);

	//获得Excel所有的Sheet名称
	std::vector<string> GetExcelSheetNames();

	//获得工作表已使用的行列数
	bool GetWorkbookRC(string inSheetName);
};

#endif

