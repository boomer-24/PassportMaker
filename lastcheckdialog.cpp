#include "lastcheckdialog.h"
#include "ui_lastcheckdialog.h"

LastCheckDialog::LastCheckDialog(QWidget *_parent) :
    QDialog(_parent), ui_(new Ui::LastCheckDialog)
{
    ui_->setupUi(this);
    this->setModal(true);
    this->setWindowTitle("Финальная проверка");
}

LastCheckDialog::LastCheckDialog(QWidget *_parent, QVector<QPair<QString, bool>> &_vFields) :
    QDialog(_parent), ui_(new Ui::LastCheckDialog)
{
    ui_->setupUi(this);
    this->setModal(true);
    this->setWindowTitle("Финальная проверка");

    this->ui_->tableWidget_lastCheck->setColumnCount(2);
    this->ui_->tableWidget_lastCheck->setRowCount(_vFields.size());
    //    QStringList tableRowNames{"№ перемычки", "Положение"};
    //    this->ui_->tableWidget_jumpers->setVerticalHeaderLabels(tableRowNames);
    this->ui_->tableWidget_lastCheck->horizontalHeader()->setSectionResizeMode(QHeaderView::ResizeToContents);

    for (int i = 0; i < _vFields.size(); i++)
    {
        //        const QPair<QString, bool>& pair = _vFields.at(i);

        //        QTableWidgetItem* item1 = new QTableWidgetItem(pair.first);
        //        this->ui_->tableWidget_lastCheck->setItem(i, 0, item1);
        //        QTableWidgetItem* item2 = new QTableWidgetItem(pair.second);
        //        this->ui_->tableWidget_lastCheck->setItem(i, 1, item2);
        //        (item2-) ? (this->ui_->tableWidget_lastCheck->item(i, 1)->setBackground(Qt::green)) :
        //                  (this->ui_->tableWidget_lastCheck->item(i, 1)->setBackground(Qt::red));
    }
}

LastCheckDialog::~LastCheckDialog()
{
    delete ui_;
}

void LastCheckDialog::FillTable(QVector<QPair<QString, QString> > &_vFields)
{
    this->ui_->tableWidget_lastCheck->setColumnCount(2);
    this->ui_->tableWidget_lastCheck->setRowCount(_vFields.size());
    this->ui_->tableWidget_lastCheck->horizontalHeader()->setSectionResizeMode(QHeaderView::ResizeToContents);

    for (int i = 0; i < _vFields.size(); i++)
    {
        const QPair<QString, QString>& pair = _vFields.at(i);

        QTableWidgetItem* item1 = new QTableWidgetItem(pair.first);
        this->ui_->tableWidget_lastCheck->setItem(i, 0, item1);
        QTableWidgetItem* item2 = new QTableWidgetItem();
        this->ui_->tableWidget_lastCheck->setItem(i, 1, item2);
        if (pair.second.isEmpty())
        {
            //            item2->setText("НЕ ЗАПОЛНЕНО");
            this->ui_->tableWidget_lastCheck->item(i, 1)->setBackground(QColor(255, 175, 200));
        } else
        {
            if (QFile::exists(pair.second))
            {
                QPixmap pixmap(pair.second);
                item2->setData(Qt::DecorationRole, pixmap);
                this->ui_->tableWidget_lastCheck->setRowHeight(item2->row(), pixmap.height());
                this->ui_->tableWidget_lastCheck->setColumnWidth(item2->column(), pixmap.width());
            } else item2->setText(pair.second);
            this->ui_->tableWidget_lastCheck->item(i, 1)->setBackground(QColor(100, 220, 120));
        }
    }
    int column0width = this->ui_->tableWidget_lastCheck->columnWidth(0);
    int column1width = this->ui_->tableWidget_lastCheck->columnWidth(1);
    int rowsHeight = 0;
    for (int i = 0; i < this->ui_->tableWidget_lastCheck->rowCount(); i++)
        rowsHeight += this->ui_->tableWidget_lastCheck->rowHeight(i);
    this->ui_->tableWidget_lastCheck->setFixedSize(column0width + column1width + 10, rowsHeight + 10);
    this->setFixedSize(QSize(this->ui_->tableWidget_lastCheck->x(), this->ui_->tableWidget_lastCheck->y()) +
                       this->ui_->tableWidget_lastCheck->size() +
                       QSize(this->ui_->tableWidget_lastCheck->x(), this->ui_->tableWidget_lastCheck->y()) +
                       QSize(0, this->ui_->pushButton_Cancel->height()));
    this->ui_->pushButton_Cancel->move(this->width() - this->ui_->pushButton_createPassport->width() * 2 - 10,
                                       this->height() - this->ui_->pushButton_Cancel->height() - 10);
    this->ui_->pushButton_createPassport->move(this->width() - this->ui_->pushButton_createPassport->width() - 10,
                                               this->height() - this->ui_->pushButton_createPassport->height() - 10);
}

void LastCheckDialog::on_pushButton_createPassport_clicked()
{
    emit this->signalCreatePassport();
}

void LastCheckDialog::on_pushButton_Cancel_clicked()
{
    this->close();
}
