#include "mainwindow.h"
#include "ui_mainwindow.h"
#include "excelOperator.h"
#include <QDebug.h>

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::on_pushButton_excel_clicked()
{
	QString macBeginStr = ui->lineEdit_begin->text();
	QString macEndStr = ui->lineEdit_end->text();

	if (macBeginStr.isEmpty() || macEndStr.isEmpty())
	{
		QMessageBox::information(this, tr("error"), tr("Please enter MAC correctly"));
		return;
	}
	QStringList macBeginList = macBeginStr.split(":", QString::SkipEmptyParts);
	macBeginStr = macBeginList.join("").trimmed();
	QStringList macEndList = macEndStr.split(":", QString::SkipEmptyParts);
	macEndStr = macEndList.join("").trimmed();
	if (macBeginStr.size() != 12 || macEndStr.size() !=12)
	{
		QMessageBox::information(this, tr("error"), tr("Please enter MAC correctly"));
		return;
	}
	
	//string to long long
	bool ok;
	qlonglong macBeginLong = macBeginStr.toLongLong(&ok, 16);
	if (!ok)
	{
		QMessageBox::information(this, tr("error"), tr("Please enter MAC correctly"));
		return;
	}
	qlonglong macEndLong = macEndStr.toLongLong(&ok, 16);
	if (!ok)
	{
		QMessageBox::information(this, tr("error"), tr("Please enter MAC correctly"));
		return;
	}
	//get file name
	QString fileName = QFileDialog::getSaveFileName(NULL, "Save File", "./MAC.xls", "Excel File(*.xls;; *.xlsx )");
	if (fileName.isEmpty())
	{
		return;
	}
	onExportExcel(fileName, macBeginLong, macEndLong);
	QMessageBox::information(this, tr("info"), tr("export excel success"));
}

void MainWindow::on_pushButton_txt_clicked()
{
	QString macBeginStr = ui->lineEdit_begin->text();
	QString macEndStr = ui->lineEdit_end->text();

	if (macBeginStr.isEmpty() || macEndStr.isEmpty())
	{
		QMessageBox::information(this, tr("error"), tr("Please enter MAC correctly"));
		return;
	}
	QStringList macBeginList = macBeginStr.split(":", QString::SkipEmptyParts);
	macBeginStr = macBeginList.join("").trimmed();
	QStringList macEndList = macEndStr.split(":", QString::SkipEmptyParts);
	macEndStr = macEndList.join("").trimmed();
	if (macBeginStr.size() != 12 || macEndStr.size() != 12)
	{
		QMessageBox::information(this, tr("error"), tr("Please enter MAC correctly"));
		return;
	}

	//string to long long
	bool ok;
	qlonglong macBeginLong = macBeginStr.toLongLong(&ok, 16);
	if (!ok)
	{
		QMessageBox::information(this, tr("error"), tr("Please enter MAC correctly"));
		return;
	}
	qlonglong macEndLong = macEndStr.toLongLong(&ok, 16);
	if (!ok)
	{
		QMessageBox::information(this, tr("error"), tr("Please enter MAC correctly"));
		return;
	}
	//get file name
	QString fileName = QFileDialog::getSaveFileName(NULL, "Save File", "./MAC.txt", "txt File(*.txt )");
	if (fileName.isEmpty())
	{
		return;
	}
	QFile file(fileName);
	if (file.exists())
	{
		file.remove();
	}
	file.open(QIODevice::WriteOnly | QIODevice::Text);
	QTextStream in(&file);
	for (qlonglong tempLong = macBeginLong; tempLong <= macEndLong; tempLong++)
	{
		in << QString::number(tempLong, 16)<<endl;
	}
	file.flush();
	file.close();
	QMessageBox::information(this, tr("info"), tr("export excel success"));
}
void MainWindow::onExportExcel(const QString &fileName, qlonglong macBeginLong, qlonglong macEndLong)
{
	ExcelOperator *excelOperator = new ExcelOperator();
	excelOperator->newExcel(fileName, true);
	excelOperator->initPSheet(1);
	int row = 1;
	for (qlonglong tempLong = macBeginLong; tempLong <= macEndLong; tempLong++)
	{
		excelOperator->setCellValue(row, 1, QString::number(tempLong,16));
		row++;
	}
	excelOperator->saveExcel(fileName);
	excelOperator->freeExcel();
}