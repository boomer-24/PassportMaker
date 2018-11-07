#ifndef CONNECTIONDIALOG_H
#define CONNECTIONDIALOG_H

#include <QApplication>
#include <QDesktopWidget>
#include <QDialog>
#include <QFile>
#include <QDebug>
#include <QTableWidgetItem>

namespace Ui {
class ConnectionDialog;
}

class ConnectionDialog : public QDialog
{
    Q_OBJECT

public:
    explicit ConnectionDialog(QWidget *_parent = 0);
    ~ConnectionDialog();
    void InsertTester(int _tester);
    void InsertPin(const QString& _pin);
    void ResizeLikeTableSize();
    bool SwapColumns(int column1, int column2);
    bool isTableWidgetContainsTester(int _tester);
    bool isTableWidgetContainsPin(const QString &_pin);
    void InsertTesterFromXML(const QString& _tester);
    void RemoveColumn(int _columnNumber);
    void RemoveRow(int _rowNumber);
    QVector<QVector<QString>> getTableContent();

private slots:
    void on_pushButton_pin1_clicked();
    void on_pushButton_pin2_clicked();
    void on_pushButton_pin3_clicked();
    void on_pushButton_pin4_clicked();
    void on_pushButton_handler_clicked();
    void on_pushButton_svpn_clicked();
    void on_pushButton_removeSelectedColumn_clicked();
    void on_pushButton_removeSelectedRow_clicked();    
    void on_pushButton_insertTester_clicked();


    void on_pushButton_user_clicked();

private:
    Ui::ConnectionDialog *ui_;
    QStringList slTesters_;
};

#endif // CONNECTIONDIALOG_H
