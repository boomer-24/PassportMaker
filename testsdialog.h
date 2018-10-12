#ifndef TESTSDIALOG_H
#define TESTSDIALOG_H

#include "test.h"
#include <QDialog>

namespace Ui {
class TestsDialog;
}

class TestsDialog : public QDialog
{
    Q_OBJECT

public:
    explicit TestsDialog(QWidget* parent = 0);
    ~TestsDialog();
    void SetTests(const QStringList& _tests);
    void SetFilters(const QStringList& _filters);
    QVector<Test> GetVerifiableTests() const;
    QVector<Test> GetUnVerifiableTests() const;
    QStringList GetTests() const;
    QStringList GetFilters() const;

private slots:
    void on_pushButton_appendVerifiableTest_clicked();
    void on_pushButton_appendUnVerifiableTest_clicked();
    void on_pushButton_inputNewTest_clicked();
    void on_pushButton_removeSelectedTest_clicked();
    void on_pushButton_removeVerifiableTestFromTable_clicked();
    void on_pushButton_removeUnVerifiableTestFromTable_clicked();
    void on_comboBox_verifiableTests_filter_currentIndexChanged(const QString &arg1);
    void on_comboBox_unverifiableTests_filter_currentIndexChanged(const QString &arg1);

private:
    Ui::TestsDialog *ui_;
    QStringList slTests_;

    // QWidget interface
protected:
    void resizeEvent(QResizeEvent *event) override;
};

#endif // TESTSDIALOG_H
