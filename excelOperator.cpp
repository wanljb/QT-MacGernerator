#include "ExcelOperator.h"

ExcelOperator::ExcelOperator()
{
}

ExcelOperator::~ExcelOperator()
{
}

void ExcelOperator::newExcel(const QString &fileName,bool isNew)
{
	pApplication = new QAxObject("Excel.Application");
	if (pApplication->isNull())
	{
		if (pApplication != NULL)//网络中很多使用excel==NULL判断，是错误的
		{
			pApplication->dynamicCall("Quit()");
			delete pApplication;
		}
		QMessageBox::critical(0, "error", "NO EXCEL APPPLICATION");
		return;
	}
	pApplication->dynamicCall("SetVisible (bool)", false);//不显示窗体
	pApplication->setProperty("DisplayAlerts", false);//不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
	pWorkBooks = pApplication->querySubObject("WorkBooks");//获取工作簿集合
	if (isNew || !QFile::exists(fileName))
	{
		pWorkBooks->dynamicCall("Add");//新建一个工作簿
		pWorkBook = pApplication->querySubObject("ActiveWorkBook");//获取当前工作簿
	}
	else 
	{
		pWorkBook = pWorkBooks->querySubObject("Open(const QString &)", fileName);
	}
	

	pSheets = pWorkBook->querySubObject("Sheets");//获取工作表集合
	pSheet = pSheets->querySubObject("Item(int)", 1);//获取工作表集合的工作表1，即sheet1
}
void ExcelOperator::initPSheet(int sheetNum)
{
	pSheet = pSheets->querySubObject("Item(int)", sheetNum);//获取工作表集合的工作表1，即sheet1
}
void ExcelOperator::appendSheet(const QString &sheetName, int sheetNum)
{
	QAxObject *pLastSheet = pSheets->querySubObject("Item(int)", sheetNum);
	pSheet = pSheets->querySubObject("Add(QVariant)", pLastSheet->asVariant());
	//pSheet = pSheets->querySubObject("Item(int)", sheetNum);
	pLastSheet->dynamicCall("Move(QVariant)", pSheet->asVariant());
	//pSheet->setProperty("Name", sheetName);
	pSheet->dynamicCall("Name", sheetName);
}

void ExcelOperator::deleteSheet(int sheetNum)
{
	QAxObject *first_sheet = pSheets->querySubObject("Item(int)", sheetNum);
	first_sheet->dynamicCall("delete");
}

void ExcelOperator::setCellValue(int row, int column, const QString &value)
{
	QAxObject *pRange = pSheet->querySubObject("Cells(int,int)", row, column);
	//pRange.setProperty("Value", value);
	pRange->dynamicCall("Value", value);

}
void ExcelOperator::readExcelData()
{
	QAxObject *cell = NULL;
	int sheet_count = pSheets->property("Count").toInt();  //获取工作表数目
	QAxObject * usedrange = pSheet->querySubObject("UsedRange");//获取该sheet的使用范围对象
	QAxObject * rows = usedrange->querySubObject("Rows");
	QAxObject * columns = usedrange->querySubObject("Columns");
	int intRowStart = usedrange->property("Row").toInt();
	int intColStart = usedrange->property("Column").toInt();
	int intRows = rows->property("Count").toInt();
	int intCols = columns->property("Count").toInt();

	// 清空数据
	for (int i = intRowStart; i <intRowStart + intRows; i++)
	{
		for (int j = intColStart; j<intColStart + intCols; j++)
		{
			cell = pSheet->querySubObject("Cells(int,int)", i, j); //获取单元格
			cell->setProperty("Value", "");
		}
	}
	// 插入数据
	//for (int i = intRowStart; i < intRowStart + rowCount; i++)
	//{
	//	for (int j = intColStart; j < intColStart + colCount; j++)
	//	{
	//		str = ui.tableWidgetExcel->item(i - intRowStart, j - intColStart)->text();
	//		cell = worksheet->querySubObject("Cells(int,int)", i, j);//获取单元格
	//		cell->setProperty("Value", str);
	//	}
	//}

}
void ExcelOperator::saveExcel(const QString &fileName)
{
	pWorkBook->dynamicCall("SaveAs(const QString &)",QDir::toNativeSeparators(fileName));
	pWorkBook->dynamicCall("Close()");//关闭工作簿
}

void ExcelOperator::freeExcel()
{
	
	if (pApplication != NULL)
	{
		pApplication->dynamicCall("Quit()");
		delete pWorkBook;
		delete pWorkBooks;
		delete pApplication;
		pWorkBook = NULL;
		pWorkBooks = NULL;
		pApplication = NULL;

	}
}