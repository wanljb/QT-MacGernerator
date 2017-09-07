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
		if (pApplication != NULL)//�����кܶ�ʹ��excel==NULL�жϣ��Ǵ����
		{
			pApplication->dynamicCall("Quit()");
			delete pApplication;
		}
		QMessageBox::critical(0, "error", "NO EXCEL APPPLICATION");
		return;
	}
	pApplication->dynamicCall("SetVisible (bool)", false);//����ʾ����
	pApplication->setProperty("DisplayAlerts", false);//����ʾ�κξ�����Ϣ�����Ϊtrue��ô�ڹر��ǻ�������ơ��ļ����޸ģ��Ƿ񱣴桱����ʾ
	pWorkBooks = pApplication->querySubObject("WorkBooks");//��ȡ����������
	if (isNew || !QFile::exists(fileName))
	{
		pWorkBooks->dynamicCall("Add");//�½�һ��������
		pWorkBook = pApplication->querySubObject("ActiveWorkBook");//��ȡ��ǰ������
	}
	else 
	{
		pWorkBook = pWorkBooks->querySubObject("Open(const QString &)", fileName);
	}
	

	pSheets = pWorkBook->querySubObject("Sheets");//��ȡ��������
	pSheet = pSheets->querySubObject("Item(int)", 1);//��ȡ�������ϵĹ�����1����sheet1
}
void ExcelOperator::initPSheet(int sheetNum)
{
	pSheet = pSheets->querySubObject("Item(int)", sheetNum);//��ȡ�������ϵĹ�����1����sheet1
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
	int sheet_count = pSheets->property("Count").toInt();  //��ȡ��������Ŀ
	QAxObject * usedrange = pSheet->querySubObject("UsedRange");//��ȡ��sheet��ʹ�÷�Χ����
	QAxObject * rows = usedrange->querySubObject("Rows");
	QAxObject * columns = usedrange->querySubObject("Columns");
	int intRowStart = usedrange->property("Row").toInt();
	int intColStart = usedrange->property("Column").toInt();
	int intRows = rows->property("Count").toInt();
	int intCols = columns->property("Count").toInt();

	// �������
	for (int i = intRowStart; i <intRowStart + intRows; i++)
	{
		for (int j = intColStart; j<intColStart + intCols; j++)
		{
			cell = pSheet->querySubObject("Cells(int,int)", i, j); //��ȡ��Ԫ��
			cell->setProperty("Value", "");
		}
	}
	// ��������
	//for (int i = intRowStart; i < intRowStart + rowCount; i++)
	//{
	//	for (int j = intColStart; j < intColStart + colCount; j++)
	//	{
	//		str = ui.tableWidgetExcel->item(i - intRowStart, j - intColStart)->text();
	//		cell = worksheet->querySubObject("Cells(int,int)", i, j);//��ȡ��Ԫ��
	//		cell->setProperty("Value", str);
	//	}
	//}

}
void ExcelOperator::saveExcel(const QString &fileName)
{
	pWorkBook->dynamicCall("SaveAs(const QString &)",QDir::toNativeSeparators(fileName));
	pWorkBook->dynamicCall("Close()");//�رչ�����
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