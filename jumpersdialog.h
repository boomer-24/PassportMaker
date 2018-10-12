#ifndef JUMPERSDIALOG_H
#define JUMPERSDIALOG_H

#include <QDialog>
#include <QDebug>

namespace Ui {
class JumpersDialog;
}

class JumpersDialog : public QDialog
{
    Q_OBJECT

public:
    explicit JumpersDialog(QWidget *parent = 0);
    ~JumpersDialog();

    void SetKY(const QString& _KY);
    void RemoveColumn(int _colomnNumber);
    QString GetJ(int _number);
    QVector<QPair<QString, QString>> GetTableContent() const;
    void ResizeLikeTableSize();

private slots:
    void on_pushButton_appendColomn_clicked();
    void on_pushButton_removeColomn_clicked();
    void on_pushButton_jumperPlaced_clicked();
    void on_pushButton_jumperNotPlaced_clicked();    
    void slotResizeLikeTableSize();

private:
    Ui::JumpersDialog *ui_;

    // QWidget interface
//protected:
//    void resizeEvent(QResizeEvent *event) override;
};

#endif // JUMPERSDIALOG_H
