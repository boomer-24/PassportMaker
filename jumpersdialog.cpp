#include "jumpersdialog.h"
#include "ui_jumpersdialog.h"

JumpersDialog::JumpersDialog(QWidget *parent) : QDialog(parent), ui_(new Ui::JumpersDialog)
{
    this->ui_->setupUi(this);
    this->setMinimumSize(330, 190);
    this->setModal(true);
    this->setWindowTitle("Установка джамперов");

    this->ui_->tableWidget_jumpers->setColumnCount(3);
    this->ui_->tableWidget_jumpers->setRowCount(2);
    QStringList tableRowNames{"№ перемычки", "Положение"};
    this->ui_->tableWidget_jumpers->setVerticalHeaderLabels(tableRowNames);
    this->ui_->tableWidget_jumpers->horizontalHeader()->setSectionResizeMode(QHeaderView::ResizeToContents);
    for (int i = 0; i < 3; i++)
    {
        QTableWidgetItem* item1 = new QTableWidgetItem(this->GetJ(i + 1));
        this->ui_->tableWidget_jumpers->setItem(0, i, item1);
        QTableWidgetItem* item2 = new QTableWidgetItem();
        this->ui_->tableWidget_jumpers->setItem(1, i, item2);
    }
    this->ui_->tableWidget_jumpers->item(1, 0)->setText("1-2");
    this->ui_->tableWidget_jumpers->item(1, 1)->setText("Установлена");
    this->ui_->tableWidget_jumpers->item(1, 2)->setText("Установлена");
    QObject::connect(this->ui_->tableWidget_jumpers, SIGNAL(cellChanged(int,int)), this, SLOT(slotResizeLikeTableSize()));
    this->ResizeLikeTableSize();
}

JumpersDialog::~JumpersDialog()
{
    delete this->ui_;
}

void JumpersDialog::RemoveColumn(int _colomnNumber)
{
    this->ui_->tableWidget_jumpers->removeColumn(_colomnNumber);
    this->ResizeLikeTableSize();
}

QString JumpersDialog::GetJ(int _number)
{
    return QString("J") + QString::number(_number);
}

QVector<QPair<QString, QString>> JumpersDialog::GetTableContent() const
{
    QVector<QPair<QString, QString>> vJPos;
    for (int i = 0; i < this->ui_->tableWidget_jumpers->columnCount(); i++)
    {
        if (this->ui_->tableWidget_jumpers->item(1, i) == nullptr)
        {
            QTableWidgetItem* item = new QTableWidgetItem();
            this->ui_->tableWidget_jumpers->setItem(1, i, item);
        }
        vJPos.push_back(QPair<QString, QString>(this->ui_->tableWidget_jumpers->item(0, i)->text(),
                                                this->ui_->tableWidget_jumpers->item(1, i)->text()));
    }
    return vJPos;
}

void JumpersDialog::ResizeLikeTableSize()
{
    int columnsWidth = 0;
    int rowsHeight = 0;
    for (int i = 0; i < this->ui_->tableWidget_jumpers->columnCount(); i++)
        columnsWidth += this->ui_->tableWidget_jumpers->columnWidth(i);
    for (int i = 0; i < this->ui_->tableWidget_jumpers->rowCount(); i++)
        rowsHeight += this->ui_->tableWidget_jumpers->rowHeight(i);
    this->ui_->tableWidget_jumpers->setFixedSize(this->ui_->tableWidget_jumpers->verticalHeader()->width() + columnsWidth + 10,
                                                 this->ui_->tableWidget_jumpers->horizontalHeader()->height() + rowsHeight + 10);
    this->setFixedSize(QSize(this->ui_->tableWidget_jumpers->x(), this->ui_->tableWidget_jumpers->y()) +
                       this->ui_->tableWidget_jumpers->size() + QSize(10, 10));
    this->setMinimumWidth(this->ui_->pushButton_jumperNotPlaced->x() + this->ui_->pushButton_jumperNotPlaced->width());
}

void JumpersDialog::on_pushButton_appendColomn_clicked()
{
    this->ui_->tableWidget_jumpers->insertColumn(this->ui_->tableWidget_jumpers->currentColumn() + 1);
    for (int i = 0; i < this->ui_->tableWidget_jumpers->columnCount(); i++)
    {
        if (this->ui_->tableWidget_jumpers->item(0, i) == nullptr)
        {
            QTableWidgetItem* item = new QTableWidgetItem();
            this->ui_->tableWidget_jumpers->setItem(0, i, item);
        }
        this->ui_->tableWidget_jumpers->item(0, i)->setText(this->GetJ(i + 1));
    }
    this->ResizeLikeTableSize();
}

void JumpersDialog::on_pushButton_removeColomn_clicked()
{
    this->ui_->tableWidget_jumpers->removeColumn(this->ui_->tableWidget_jumpers->currentColumn());
    for (int i = 0; i < this->ui_->tableWidget_jumpers->columnCount(); i++)
    {
        if (this->ui_->tableWidget_jumpers->item(0, i) == nullptr)
        {
            QTableWidgetItem* item = new QTableWidgetItem();
            this->ui_->tableWidget_jumpers->setItem(0, i, item);
        }
        this->ui_->tableWidget_jumpers->item(0, i)->setText(this->GetJ(i + 1));
    }
    this->ResizeLikeTableSize();
}

void JumpersDialog::on_pushButton_jumperPlaced_clicked()
{
    QTableWidgetItem* item = new QTableWidgetItem();
    item->setText("Установлена");
    this->ui_->tableWidget_jumpers->setItem(this->ui_->tableWidget_jumpers->currentRow(),
                                            this->ui_->tableWidget_jumpers->currentColumn(), item);
    this->ResizeLikeTableSize();
}

void JumpersDialog::on_pushButton_jumperNotPlaced_clicked()
{
    QTableWidgetItem* item = new QTableWidgetItem();
    item->setText("Не установлена");
    this->ui_->tableWidget_jumpers->setItem(this->ui_->tableWidget_jumpers->currentRow(),
                                            this->ui_->tableWidget_jumpers->currentColumn(), item);
    this->ResizeLikeTableSize();
}

void JumpersDialog::slotResizeLikeTableSize()
{
    this->ResizeLikeTableSize();
}
