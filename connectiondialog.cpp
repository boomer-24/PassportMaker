#include "connectiondialog.h"
#include "ui_connectiondialog.h"

ConnectionDialog::ConnectionDialog(QWidget *_parent) : QDialog(_parent), ui_(new Ui::ConnectionDialog)
{
    ui_->setupUi(this);
    this->setModal(true);
    this->setWindowTitle("Подключение к тестеру");
    this->ui_->tableWidget_connection->insertRow(0);
    this->ui_->tableWidget_connection->insertColumn(0);
    QTableWidgetItem *item00 = new QTableWidgetItem();
    this->ui_->tableWidget_connection->setItem(0, 0, item00);
    QString slashPath(QCoreApplication::applicationDirPath().append("/slash.png"));
    if (QFile::exists(slashPath))
    {
        QPixmap pixmap(slashPath);
        item00->setData(Qt::DecorationRole, pixmap);
        this->ui_->tableWidget_connection->setRowHeight(item00->row(), pixmap.height());
        this->ui_->tableWidget_connection->setColumnWidth(item00->column(), pixmap.width());
    } else qDebug() << slashPath << " not exist";
    this->ResizeLikeTableSize();
}

ConnectionDialog::~ConnectionDialog()
{
    delete ui_;
}

void ConnectionDialog::InsertTester(int _tester)
{
    if (this->ui_->tableWidget_connection->columnCount() == 1)
    {
        this->ui_->tableWidget_connection->insertColumn(1);
        QTableWidgetItem* item = new QTableWidgetItem(QString::number(_tester));
        this->ui_->tableWidget_connection->setItem(0, 1, item);
        for (int row = 1; row < this->ui_->tableWidget_connection->rowCount(); row++)
            this->ui_->tableWidget_connection->setItem(row, 1, new QTableWidgetItem(""));
    } else
    {
        int count = this->ui_->tableWidget_connection->columnCount();
        for (int i = 1; i < count; i++)
        {
            int currentColumn = this->ui_->tableWidget_connection->item(0, i)->text().toInt();
            if (_tester <= currentColumn)
            {
                this->ui_->tableWidget_connection->insertColumn(i);
                QTableWidgetItem* item = new QTableWidgetItem(QString::number(_tester));
                this->ui_->tableWidget_connection->setItem(0, i, item);
                for (int row = 1; row < this->ui_->tableWidget_connection->rowCount(); row++)
                    this->ui_->tableWidget_connection->setItem(row, i, new QTableWidgetItem(""));
                break;
            }
            if (currentColumn < _tester)
            {
                this->ui_->tableWidget_connection->insertColumn(i + 1);
                QTableWidgetItem* item = new QTableWidgetItem(QString::number(_tester));
                this->ui_->tableWidget_connection->setItem(0, i + 1, item);
                for (int row = 1; row < this->ui_->tableWidget_connection->rowCount(); row++)
                    this->ui_->tableWidget_connection->setItem(row, i + 1, new QTableWidgetItem(""));
                break;
            }
        }
    }
    this->ResizeLikeTableSize();
}

void ConnectionDialog::ResizeLikeTableSize()
{
    int columnsWidth = 0;
    int rowsHeight = 0;
    for (int i = 0; i < this->ui_->tableWidget_connection->columnCount(); i++)
        columnsWidth += this->ui_->tableWidget_connection->columnWidth(i);
    for (int i = 0; i < this->ui_->tableWidget_connection->rowCount(); i++)
        rowsHeight += this->ui_->tableWidget_connection->rowHeight(i);
    this->ui_->tableWidget_connection->setFixedSize(this->ui_->tableWidget_connection->verticalHeader()->width() + columnsWidth + 10,
                                                    this->ui_->tableWidget_connection->horizontalHeader()->height() + rowsHeight + 10);
    this->setFixedSize(QSize(this->ui_->tableWidget_connection->x(), this->ui_->tableWidget_connection->y()) +
                       this->ui_->tableWidget_connection->size() + QSize(100, 500));
    this->setMinimumWidth(this->ui_->pushButton->x() + this->ui_->pushButton->width());
}

bool ConnectionDialog::SwapColumns(int column1, int column2)
{
    int rowCount(this->ui_->tableWidget_connection->rowCount());
    int columnCount(this->ui_->tableWidget_connection->columnCount());
    if ((0 < column1 && column1 < columnCount) && (0 < column2 && column2 < columnCount))
    {
        for (int row = 0; row < rowCount; row++)
        {
            QString val1("");
            if (this->ui_->tableWidget_connection->item(row, column1) != nullptr)
                val1 = this->ui_->tableWidget_connection->item(row, column1)->text() ;
            else
                this->ui_->tableWidget_connection->setItem(row, column1, new QTableWidgetItem(""));
            QString val2("");
            if (this->ui_->tableWidget_connection->item(row, column2) != nullptr)
                val2 = this->ui_->tableWidget_connection->item(row, column2)->text();
            else
                this->ui_->tableWidget_connection->setItem(row, column2, new QTableWidgetItem(""));
            this->ui_->tableWidget_connection->item(row, column1)->setText(val2);
            this->ui_->tableWidget_connection->item(row, column2)->setText(val1);
        }
    } else
        return false;
}

void ConnectionDialog::on_pushButton_clicked()
{
    if (!this->ui_->lineEdit->text().isEmpty())
        this->InsertTester(this->ui_->lineEdit->text().toInt());
    this->ui_->lineEdit->clear();
    this->ui_->tableWidget_connection->insertRow(1);
}

void ConnectionDialog::on_pushButton_2_clicked()
{
    qDebug() << this->SwapColumns(1, 2);
}
