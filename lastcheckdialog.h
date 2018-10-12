#ifndef LASTCHECKDIALOG_H
#define LASTCHECKDIALOG_H

#include <QDialog>
#include <QFile>
#include <QDebug>

namespace Ui {
class LastCheckDialog;
}

class LastCheckDialog : public QDialog
{
    Q_OBJECT

public:
    explicit LastCheckDialog(QWidget *_parent = 0);
    LastCheckDialog(QWidget *_parent, QVector<QPair<QString, bool>>& _vFields);
    ~LastCheckDialog();
    void FillTable(QVector<QPair<QString, QString> > &_vFields);

signals:
    void signalCreatePassport();

private slots:
    void on_pushButton_createPassport_clicked();
    void on_pushButton_Cancel_clicked();

private:
    Ui::LastCheckDialog *ui_;

    // QWidget interface
//protected:
//    void resizeEvent(QResizeEvent *event) override;
};

#endif // LASTCHECKDIALOG_H
