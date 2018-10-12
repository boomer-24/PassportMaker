#ifndef CONNECTIONDIALOG_H
#define CONNECTIONDIALOG_H

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
    void ResizeLikeTableSize();
    bool SwapColumns(int column1, int column2);

private slots:
    void on_pushButton_clicked();

    void on_pushButton_2_clicked();

private:
    Ui::ConnectionDialog *ui_;
};

#endif // CONNECTIONDIALOG_H
