#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include "passportmaker2K.h"
#include "passportmakertt.h"
#include "testsdialog.h"
#include "jumpersdialog.h"
#include "connectiondialog.h"
#include "lastcheckdialog.h"

#include <QMainWindow>
#include <QFileDialog>
#include <QProcess>
#include <QStringList>
#include <QDesktopServices>
#include <QXmlStreamWriter>
#include <QCursor>
#include <QMessageBox>
#include <QPair>
#include <QThread>

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT
    Ui::MainWindow *ui_;
    TestsDialog testDialog_;
    JumpersDialog jumpersDialog_;
    ConnectionDialog connectionDialog_;
    LastCheckDialog lastCheckDialog2K_;
    LastCheckDialog lastCheckDialogTT_;
    QThread thread2K_, threadTT_;

    QString dirPath_2K_, dirPath_TT_;
    QStringList allAnimals_, analogAnimals_, digitAnimals_, animalsPositions_;
    QString analogChief_, digitChief_;
    QStringList slElements2K_, slElementsTT_, slFormats2K_, slFormatsTT_;
    QStringList slTesters2K_;
    QString commentConnectionKYTT_, adapterName2K_, adapterNumber2K_;

public:
    explicit MainWindow(QWidget *_parent = 0);
    ~MainWindow();
    bool isOnlySpace(const QString& _str);
    void Initialize(const QString& _path);
    bool isWordRunning();
    void InitializeTestsTT(const QString& _path);
    QDomElement MakeElement(QDomDocument& _doc, const QString& _name, const QString& _value);
    void SaveTestXml(const QStringList &_slTests, const QStringList &_slFilters);    

protected:
    void resizeEvent(QResizeEvent *event) override;

public slots:
    void slotCreatePassport2K();
    void slotCreatePassportTT();

private slots:
    void on_comboBox_devName_2k_activated(const QString &arg1);
    void on_comboBox_devPosition_2k_activated(const QString &arg1);
    void on_comboBox_devName_tt_activated(const QString &arg1);
    void on_comboBox_devPosition_tt_activated(const QString &arg1);

    void on_pushButton_addElement_2k_clicked();
    void on_pushButton_removeElement_2k_clicked();
    void on_pushButton_addFormat_2k_clicked();
    void on_pushButton_removeFormat_2k_clicked();
    void on_pushButton_selectDirectory_2k_clicked();
    void on_pushButton_createSavePasport_2k_clicked();

    void on_pushButton_addElement_tt_clicked();
    void on_pushButton_removeElement_tt_clicked();
    void on_pushButton_addFormat_tt_clicked();
    void on_pushButton_removeFormat_tt_clicked();
    void on_pushButton_selectDirectory_tt_clicked();
    void on_pushButton_openPaint_tt_clicked();
    void on_pushButton_createSavePasport_tt_clicked();
    void on_pushButton_testDialogBox_clicked();
    void on_pushButton_openDirTT_clicked();
    void on_pushButton_openDir2K_clicked();

    void on_pushButton_jumpers_clicked();
    void on_radioButton_jumpersYes_clicked();
    void on_radioButton_jumpersNo_clicked();
    void on_pushButton_connectionDialog_clicked();

};

#endif // MAINWINDOW_H
