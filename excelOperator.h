#ifndef EXCEL_OPERATOR_H
#define EXCEL_OPERATOR_H

#include <ActiveQt/QAxObject>
#include <QFile>
#include <QDir>
#include <QMessageBox>
#include <QFileDialog>
class ExcelOperator
{
public:
	ExcelOperator();
	~ExcelOperator();
	void newExcel(const QString &fileName,bool isNew); //新建一个excel
	void initPSheet(int sheetNum);
	void appendSheet(const QString &sheetName,int sheetNum);//增加1个Worksheet
	void deleteSheet(int sheetNum);//删除worksheet
	void setCellValue(int row, int column, const QString &value);//向Excel单元格中写入数据
	void readExcelData();
	void saveExcel(const QString &fileName);//保存excel
	void freeExcel();//释放excel
	QAxObject *pApplication = NULL;
	QAxObject *pWorkBooks = NULL;
	QAxObject *pWorkBook = NULL;
	QAxObject *pSheets = NULL;
	QAxObject *pSheet = NULL;

};









#endif //EXCEL_OPERATOR_H