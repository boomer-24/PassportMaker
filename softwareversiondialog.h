#ifndef SOFTWAREVERSIONDIALOG_H
#define SOFTWAREVERSIONDIALOG_H

#include <QDialog>
#include <QTableWidgetItem>

namespace Ui {
class SoftwareVersionDialog;
}

class SoftwareVersionDialog : public QDialog
{
    Q_OBJECT

public:
    explicit SoftwareVersionDialog(QWidget *parent = 0);
    ~SoftwareVersionDialog();

    void setSlVersions(const QStringList &slVersions);
    bool isTableWidgetContainsVersion(const QString& _version) const;
    void InsertVersion(const QString& _version);
    void RemoveRow(int _rowNumber);
    void ResizeLikeTableSize();
    QString getSelectedVersion() const;
    QStringList getSlVersions() const;

private slots:
    void on_pushButton_addNewVersion_clicked();
    void on_pushButton_removeSelectedVersion_clicked();
    void slotItemChanged(int, int);
    void slotItemClicked(QTableWidgetItem* _item);

signals:
    void signalVersionValueChange(QString);

private:
    Ui::SoftwareVersionDialog *ui_;
};

#endif // SOFTWAREVERSIONDIALOG_H
