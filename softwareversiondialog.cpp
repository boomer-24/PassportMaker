#include "softwareversiondialog.h"
#include "ui_softwareversiondialog.h"

SoftwareVersionDialog::SoftwareVersionDialog(QWidget *parent) :
    QDialog(parent),
    ui_(new Ui::SoftwareVersionDialog)
{
    this->ui_->setupUi(this);
    this->setMinimumSize(270, 190);
    this->setModal(true);
    this->setWindowTitle("Версия оболочки Formula2K");

    this->ui_->tableWidget_softwareVersion->horizontalHeader()->setSectionResizeMode(QHeaderView::ResizeToContents);

    this->ui_->tableWidget_softwareVersion->setColumnCount(1);
    QObject::connect(this->ui_->tableWidget_softwareVersion, SIGNAL(cellChanged(int,int)),
                     this, SLOT(slotItemChanged(int,int)));
    QObject::connect(this->ui_->tableWidget_softwareVersion, SIGNAL(itemClicked(QTableWidgetItem*)),
                     this, SLOT(slotItemClicked(QTableWidgetItem*)));
    this->ResizeLikeTableSize();
}

SoftwareVersionDialog::~SoftwareVersionDialog()
{
    delete ui_;
}

void SoftwareVersionDialog::setSlVersions(const QStringList &slVersions)
{
    for (const QString& version : slVersions)
    {
        this->InsertVersion(version);
    }
    this->ui_->tableWidget_softwareVersion->selectRow(this->ui_->tableWidget_softwareVersion->rowCount() - 1);
}

bool SoftwareVersionDialog::isTableWidgetContainsVersion(const QString &_version) const
{
    int rowCount = this->ui_->tableWidget_softwareVersion->rowCount();
    for (int i = 0; i < rowCount; i++)
    {
        QString currentItem = this->ui_->tableWidget_softwareVersion->item(i, 0)->text();
        if (currentItem == _version)
            return true;
    }
    return false;
}

void SoftwareVersionDialog::InsertVersion(const QString &_version)
{
    if (!this->isTableWidgetContainsVersion(_version))
    {
        this->ui_->tableWidget_softwareVersion->insertRow(this->ui_->tableWidget_softwareVersion->rowCount());
        QTableWidgetItem* item = new QTableWidgetItem(_version);
        this->ui_->tableWidget_softwareVersion->setItem(this->ui_->tableWidget_softwareVersion->rowCount() - 1, 0, item);
        item->setTextAlignment(Qt::AlignHCenter);
    }
    this->ResizeLikeTableSize();
}

void SoftwareVersionDialog::RemoveRow(int _rowNumber)
{
    this->ui_->tableWidget_softwareVersion->removeRow(_rowNumber);
    this->ui_->tableWidget_softwareVersion->selectRow(this->ui_->tableWidget_softwareVersion->rowCount() - 1);
    this->ResizeLikeTableSize();
}

void SoftwareVersionDialog::ResizeLikeTableSize()
{
    int columnsWidth = 0;
    int rowsHeight = 0;
    for (int i = 0; i < this->ui_->tableWidget_softwareVersion->columnCount(); i ++)
        columnsWidth += this->ui_->tableWidget_softwareVersion->columnWidth(i);
    for (int i = 0; i < this->ui_->tableWidget_softwareVersion->rowCount(); i++)
        rowsHeight += this->ui_->tableWidget_softwareVersion->rowHeight(i);
    this->ui_->tableWidget_softwareVersion->setFixedSize(this->ui_->tableWidget_softwareVersion->verticalHeader()->width() +
                                                         columnsWidth + 10,
                                                         this->ui_->tableWidget_softwareVersion->horizontalHeader()->height() +
                                                         rowsHeight + 10);

    this->setFixedSize(QSize(this->ui_->tableWidget_softwareVersion->x(), this->ui_->tableWidget_softwareVersion->y()) +
                       this->ui_->tableWidget_softwareVersion->size() + QSize(10, 10));
    if (this->width() < 270) this->setFixedWidth(270);
    if (this->height() < this->ui_->pushButton_removeSelectedVersion->pos().y() +
            this->ui_->pushButton_removeSelectedVersion->height())
        this->setFixedHeight(this->ui_->pushButton_removeSelectedVersion->pos().y() +
                             this->ui_->pushButton_removeSelectedVersion->height());
    if (this->ui_->tableWidget_softwareVersion->width() < this->width() - this->ui_->pushButton_addNewVersion->width() - 30)
        this->ui_->tableWidget_softwareVersion->setFixedWidth(this->width()-this->ui_->pushButton_addNewVersion->width() - 30);
}

QString SoftwareVersionDialog::getSelectedVersion() const
{
    QString selectedVersion(this->ui_->tableWidget_softwareVersion->currentItem()->text());
    return selectedVersion;
}

QStringList SoftwareVersionDialog::getSlVersions() const
{
    QStringList slVersions;
    for (int r = 0; r < this->ui_->tableWidget_softwareVersion->rowCount(); r++)
    {
        QString val(this->ui_->tableWidget_softwareVersion->item(r, 0)->text());
        slVersions.push_back(val);
    }
    return slVersions;
}

void SoftwareVersionDialog::on_pushButton_addNewVersion_clicked()
{
    this->InsertVersion(QString(""));
}

void SoftwareVersionDialog::on_pushButton_removeSelectedVersion_clicked()
{
    this->RemoveRow(this->ui_->tableWidget_softwareVersion->currentRow());
    if (this->ui_->tableWidget_softwareVersion->rowCount() > 0)
        this->ui_->tableWidget_softwareVersion->selectRow(this->ui_->tableWidget_softwareVersion->rowCount() - 1);
}

void SoftwareVersionDialog::slotItemChanged(int, int)
{
    this->ResizeLikeTableSize();
}

void SoftwareVersionDialog::slotItemClicked(QTableWidgetItem *_item)
{
    QString itemValue = _item->text();
    emit this->signalVersionValueChange(itemValue);
}
