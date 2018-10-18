#include "connectiondialog.h"
#include "ui_connectiondialog.h"

ConnectionDialog::ConnectionDialog(QWidget *_parent) : QDialog(_parent), ui_(new Ui::ConnectionDialog)
{
    ui_->setupUi(this);
    this->setMaximumSize(QApplication::desktop()->size() - QSize(300, 300));
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
    this->setMinimumSize(this->ui_->pushButton_removeSelectedRow->pos().rx() +
                         this->ui_->pushButton_removeSelectedRow->width() + 15,
                         this->ui_->pushButton_svpn->pos().ry() +
                         this->ui_->pushButton_svpn->height() + 15);
    this->ResizeLikeTableSize();
}

ConnectionDialog::~ConnectionDialog()
{
    delete ui_;
}

void ConnectionDialog::InsertTester(int _tester)
{
    if (!this->isTableWidgetContainsTester(_tester))
    {
        if (this->ui_->tableWidget_connection->columnCount() == 1)
        {
            this->ui_->tableWidget_connection->insertColumn(1);
            QTableWidgetItem* item = new QTableWidgetItem(QString::number(_tester));
            this->ui_->tableWidget_connection->setItem(0, 1, item);
            item->setTextAlignment(Qt::AlignCenter);
            for (int row = 1; row < this->ui_->tableWidget_connection->rowCount(); row++)
                this->ui_->tableWidget_connection->setItem(row, 1, new QTableWidgetItem(""));
        } else
        {
            int columnCount = this->ui_->tableWidget_connection->columnCount();
            for (int i = 1; i < columnCount; i++)
            {
                int currentColumn = this->ui_->tableWidget_connection->item(0, i)->text().toInt();
                if (_tester <= currentColumn)
                {
                    this->ui_->tableWidget_connection->insertColumn(i);
                    QTableWidgetItem* item = new QTableWidgetItem(QString::number(_tester));
                    this->ui_->tableWidget_connection->setItem(0, i, item);
                    item->setTextAlignment(Qt::AlignCenter);
                    for (int row = 1; row < this->ui_->tableWidget_connection->rowCount(); row++)
                        this->ui_->tableWidget_connection->setItem(row, i, new QTableWidgetItem(""));
                    break;
                }
                if (i == columnCount - 1)
                {
                    if (currentColumn < _tester)
                    {
                        this->ui_->tableWidget_connection->insertColumn(i + 1);
                        QTableWidgetItem* item = new QTableWidgetItem(QString::number(_tester));
                        this->ui_->tableWidget_connection->setItem(0, i + 1, item);
                        item->setTextAlignment(Qt::AlignCenter);
                        for (int row = 1; row < this->ui_->tableWidget_connection->rowCount(); row++)
                            this->ui_->tableWidget_connection->setItem(row, i + 1, new QTableWidgetItem(""));
                        break;
                    }
                }
                if (i < columnCount - 1)
                {
                    if ((currentColumn < _tester) &&
                            (_tester <= this->ui_->tableWidget_connection->item(0, i + 1)->text().toInt()))
                    {
                        this->ui_->tableWidget_connection->insertColumn(i + 1);
                        QTableWidgetItem* item = new QTableWidgetItem(QString::number(_tester));
                        this->ui_->tableWidget_connection->setItem(0, i + 1, item);
                        item->setTextAlignment(Qt::AlignCenter);
                        for (int row = 1; row < this->ui_->tableWidget_connection->rowCount(); row++)
                            this->ui_->tableWidget_connection->setItem(row, i + 1, new QTableWidgetItem(""));
                        break;
                    }
                }
            }
        }
        this->ResizeLikeTableSize();
    }
}

void ConnectionDialog::InsertPin(const QString &_pin)
{
    if (!this->isTableWidgetContainsPin(_pin))
    {
        this->ui_->tableWidget_connection->insertRow(this->ui_->tableWidget_connection->rowCount());
        QTableWidgetItem* item = new QTableWidgetItem(_pin);
        this->ui_->tableWidget_connection->setItem(this->ui_->tableWidget_connection->rowCount() - 1, 0, item);
        item->setTextAlignment(Qt::AlignCenter);
        for (int c = 1; c < this->ui_->tableWidget_connection->columnCount(); c++)
        {
            QTableWidgetItem* emptyItem = new QTableWidgetItem("");
            this->ui_->tableWidget_connection->setItem(this->ui_->tableWidget_connection->rowCount() - 1, c, emptyItem);
            emptyItem->setTextAlignment(Qt::AlignCenter);
        }
        this->ResizeLikeTableSize();
    }
}

void ConnectionDialog::ResizeLikeTableSize()
{
    int columnsWidth = 0;
    int rowsHeight = 0;
    for (int i = 0; i < this->ui_->tableWidget_connection->columnCount(); i++)
        columnsWidth += this->ui_->tableWidget_connection->columnWidth(i);
    for (int i = 0; i < this->ui_->tableWidget_connection->rowCount(); i++)
        rowsHeight += this->ui_->tableWidget_connection->rowHeight(i);
    this->ui_->tableWidget_connection->setFixedSize(columnsWidth + 15, rowsHeight + 15);

//    this->setMinimumSize(this->ui_->pushButton_removeSelectedRow->pos().rx() +
//                         this->ui_->pushButton_removeSelectedRow->width() + 15,
//                         this->ui_->pushButton_svpn->pos().ry() +
//                         this->ui_->pushButton_svpn->height() + 15);
    this->setFixedSize(QSize(this->ui_->tableWidget_connection->x(), this->ui_->tableWidget_connection->y()) +
                       this->ui_->tableWidget_connection->size() + QSize(150, 150));

    //    this->setMaximumWidth(QApplication::desktop()->screenGeometry().width() - 300);
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

bool ConnectionDialog::isTableWidgetContainsTester(int _tester)
{
    int columnCount = this->ui_->tableWidget_connection->columnCount();
    for (int i = 1; i < columnCount; i++)
    {
        int currentColumn = this->ui_->tableWidget_connection->item(0, i)->text().toInt();
        if (currentColumn == _tester)
            return true;
    }
    return false;
}

bool ConnectionDialog::isTableWidgetContainsPin(const QString& _pin)
{
    int rowCount = this->ui_->tableWidget_connection->rowCount();
    for (int i = 1; i < rowCount; i++)
    {
        QString currentRow = this->ui_->tableWidget_connection->item(i, 0)->text();
        if (currentRow == _pin)
            return true;
    }
    return false;
}

void ConnectionDialog::InsertTesterFromXML(const QString &_tester)
{
    this->slTesters_.push_back(_tester);
    this->InsertTester(_tester.toInt());
    this->ResizeLikeTableSize();
}

void ConnectionDialog::RemoveColumn(int _columnNumber)
{
    if (_columnNumber != 0)
    {
        this->ui_->tableWidget_connection->removeColumn(_columnNumber);
        this->ResizeLikeTableSize();
    }
}

void ConnectionDialog::RemoveRow(int _rowNumber)
{
    if (_rowNumber != 0)
    {
        this->ui_->tableWidget_connection->removeRow(_rowNumber);
        this->ResizeLikeTableSize();
    }
}

QVector<QVector<QString>> ConnectionDialog::getTableContent()
{
    QVector<QVector<QString>> vvContent;
    for (int r = 0; r < this->ui_->tableWidget_connection->rowCount(); r++)
    {
        QVector<QString> vRow;
        for (int c = 0; c < this->ui_->tableWidget_connection->columnCount(); c++)
        {
            QString val(this->ui_->tableWidget_connection->item(r, c)->text());
            vRow.push_back(val);
        }
        vvContent.push_back(vRow);
    }
    return vvContent;
}

void ConnectionDialog::on_pushButton_pin1_clicked()
{
    if (!this->isTableWidgetContainsPin("Pin1"))
        this->InsertPin("Pin1");
}

void ConnectionDialog::on_pushButton_pin2_clicked()
{
    if (!this->isTableWidgetContainsPin("Pin2"))
        this->InsertPin("Pin2");
}

void ConnectionDialog::on_pushButton_pin3_clicked()
{
    if (!this->isTableWidgetContainsPin("Pin3"))
        this->InsertPin("Pin3");
}

void ConnectionDialog::on_pushButton_pin4_clicked()
{
    if (!this->isTableWidgetContainsPin("Pin4"))
        this->InsertPin("Pin4");
}

void ConnectionDialog::on_pushButton_svpn_clicked()
{
    if (!this->isTableWidgetContainsPin("SVPN"))
        this->InsertPin("SVPN");
    else if (!this->isTableWidgetContainsPin("SVPN2"))
        this->InsertPin("SVPN2");
}

void ConnectionDialog::on_pushButton_removeSelectedColumn_clicked()
{
    this->RemoveColumn(this->ui_->tableWidget_connection->currentColumn());
}

void ConnectionDialog::on_pushButton_removeSelectedRow_clicked()
{
    this->RemoveRow(this->ui_->tableWidget_connection->currentRow());
}

void ConnectionDialog::on_pushButton_insertTester_clicked()
{
    int tester = this->ui_->lineEdit_insertTester->text().toInt();
    if (tester)
        this->InsertTester(tester);
}
