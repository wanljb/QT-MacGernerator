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
	void newExcel(const QString &fileName,bool isNew); //�½�һ��excel
	void initPSheet(int sheetNum);
	void appendSheet(const QString &sheetName,int sheetNum);//����1��Worksheet
	void deleteSheet(int sheetNum);//ɾ��worksheet
	void setCellValue(int row, int column, const QString &value);//��Excel��Ԫ����д������
	void readExcelData();
	void saveExcel(const QString &fileName);//����excel
	void freeExcel();//�ͷ�excel
	QAxObject *pApplication = NULL;
	QAxObject *pWorkBooks = NULL;
	QAxObject *pWorkBook = NULL;
	QAxObject *pSheets = NULL;
	QAxObject *pSheet = NULL;

};









#endif //EXCEL_OPERATOR_H