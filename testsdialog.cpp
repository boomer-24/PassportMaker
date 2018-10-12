#include "testsdialog.h"
#include "ui_testsdialog.h"

TestsDialog::TestsDialog(QWidget *parent) : QDialog(parent), ui_(new Ui::TestsDialog)
{
    this->ui_->setupUi(this);
    this->setModal(true);
    this->setMinimumSize(760, 600);
    this->setWindowTitle("Добавление тестов");

    this->ui_->tableWidget_verifiableTests->setColumnCount(6);
    this->ui_->tableWidget_verifiableTests->setRowCount(0);
    QStringList tableverifiebleTestsColumnsName{"Наименование параметра,\n единица измерения", "Обозначение",
                                                "Не менее", "Не более", "Условия", "Отклонения от ТУ"};
    this->ui_->tableWidget_verifiableTests->setHorizontalHeaderLabels(tableverifiebleTestsColumnsName);
    this->ui_->tableWidget_verifiableTests->horizontalHeader()->setSectionResizeMode(QHeaderView::ResizeToContents);



    this->ui_->tableWidget_unverifiableTests->setColumnCount(5);
    this->ui_->tableWidget_unverifiableTests->setRowCount(0);
    QStringList tableUnVerifiebleTestsColumnsName{"Наименование параметра,\n единица измерения", "Обозначение",
                                                  "Не менее", "Не более", "Условия"};
    this->ui_->tableWidget_unverifiableTests->setHorizontalHeaderLabels(tableUnVerifiebleTestsColumnsName);
    this->ui_->tableWidget_unverifiableTests->horizontalHeader()->setSectionResizeMode(QHeaderView::ResizeToContents);
}

TestsDialog::~TestsDialog()
{
    delete this->ui_;
}

void TestsDialog::SetTests(const QStringList &_tests)
{
    this->slTests_ = _tests;
    this->slTests_.sort();
    //    this->ui_->comboBox_verifiableTests->addItem("");
    this->ui_->comboBox_verifiableTests->addItems(this->slTests_);
    //    this->ui_->comboBox_unverifiableTests->addItem("");
    this->ui_->comboBox_unverifiableTests->addItems(this->slTests_);
}

void TestsDialog::SetFilters(const QStringList &_filters)
{
    this->ui_->comboBox_verifiableTests_filter->addItem("Все тесты");
    this->ui_->comboBox_verifiableTests_filter->addItems(_filters);
    this->ui_->comboBox_unverifiableTests_filter->addItem("Все тесты");
    this->ui_->comboBox_unverifiableTests_filter->addItems(_filters);
}

void TestsDialog::on_pushButton_appendVerifiableTest_clicked()
{
    int rowCount = this->ui_->tableWidget_verifiableTests->rowCount();
    this->ui_->tableWidget_verifiableTests->setRowCount(rowCount + 1);
    QTableWidgetItem* itemTestName = new QTableWidgetItem();
    itemTestName->setText(this->ui_->comboBox_verifiableTests->currentText());
    this->ui_->tableWidget_verifiableTests->setItem(rowCount, 0, itemTestName);

    for (int j = 1; j < this->ui_->tableWidget_verifiableTests->columnCount(); j++)
    {
        QTableWidgetItem* item = new QTableWidgetItem();
        this->ui_->tableWidget_verifiableTests->setItem(rowCount, j, item);
        item->setText(" ");
    }
}

void TestsDialog::on_pushButton_appendUnVerifiableTest_clicked()
{
    int rowCount = this->ui_->tableWidget_unverifiableTests->rowCount();
    this->ui_->tableWidget_unverifiableTests->setRowCount(rowCount + 1);
    QTableWidgetItem* itemTestName = new QTableWidgetItem();
    itemTestName->setText(this->ui_->comboBox_unverifiableTests->currentText());
    this->ui_->tableWidget_unverifiableTests->setItem(rowCount, 0, itemTestName);

    for (int j = 1; j < this->ui_->tableWidget_unverifiableTests->columnCount(); j++)
    {
        QTableWidgetItem* item = new QTableWidgetItem();
        this->ui_->tableWidget_unverifiableTests->setItem(rowCount, j, item);
        item->setText(" ");
    }
}

QVector<Test> TestsDialog::GetVerifiableTests() const
{
    QVector<Test> vTests;
    for (int row = 0; row < this->ui_->tableWidget_verifiableTests->rowCount(); row++)
    {
        Test test(this->ui_->tableWidget_verifiableTests->item(row, 0)->text(),
                  this->ui_->tableWidget_verifiableTests->item(row, 1)->text(),
                  this->ui_->tableWidget_verifiableTests->item(row, 2)->text(),
                  this->ui_->tableWidget_verifiableTests->item(row, 3)->text(),
                  this->ui_->tableWidget_verifiableTests->item(row, 4)->text(),
                  this->ui_->tableWidget_verifiableTests->item(row, 5)->text());
        vTests.push_back(test);
    }
    if (vTests.isEmpty())
        vTests.push_back(Test(" ", " ", " ", " ", " ", " "));
    return vTests;
}

QVector<Test> TestsDialog::GetUnVerifiableTests() const
{
    QVector<Test> vUnTests;
    for (int row = 0; row < this->ui_->tableWidget_unverifiableTests->rowCount(); row++)
    {
        Test unTest(this->ui_->tableWidget_unverifiableTests->item(row, 0)->text(),
                    this->ui_->tableWidget_unverifiableTests->item(row, 1)->text(),
                    this->ui_->tableWidget_unverifiableTests->item(row, 2)->text(),
                    this->ui_->tableWidget_unverifiableTests->item(row, 3)->text(),
                    this->ui_->tableWidget_unverifiableTests->item(row, 4)->text());
        vUnTests.push_back(unTest);
    }
    if (vUnTests.isEmpty())
        vUnTests.push_back(Test(" ", " ", " ", " ", " "));
    return vUnTests;
}

QStringList TestsDialog::GetTests() const
{
    return this->slTests_;
}

QStringList TestsDialog::GetFilters() const
{
    QStringList slFilters;
    for (int i = 0; i < this->ui_->comboBox_verifiableTests_filter->count(); i++)
        if (!this->ui_->comboBox_verifiableTests_filter->itemText(i).contains("Все тесты"))
            slFilters.push_back(this->ui_->comboBox_verifiableTests_filter->itemText(i));
    return slFilters;
}

void TestsDialog::on_pushButton_inputNewTest_clicked()
{
    QString test(this->ui_->lineEdit_inputNewTest->text());
    if (this->ui_->comboBox_verifiableTests->findText(test) == -1)
    {
        this->slTests_.push_back(test);
        this->slTests_.sort();
        this->ui_->comboBox_verifiableTests->addItem(this->ui_->lineEdit_inputNewTest->text());
        this->ui_->comboBox_unverifiableTests->addItem(this->ui_->lineEdit_inputNewTest->text());
    }
}

void TestsDialog::on_pushButton_removeSelectedTest_clicked()
{
    int i = this->ui_->comboBox_verifiableTests->currentIndex();
    this->slTests_.removeOne(this->ui_->comboBox_verifiableTests->currentText());
    this->ui_->comboBox_verifiableTests->removeItem(i);
    this->ui_->comboBox_unverifiableTests->removeItem(i);
}

void TestsDialog::on_pushButton_removeVerifiableTestFromTable_clicked()
{
    this->ui_->tableWidget_verifiableTests->removeRow(this->ui_->tableWidget_verifiableTests->currentRow());
}

void TestsDialog::on_pushButton_removeUnVerifiableTestFromTable_clicked()
{
    this->ui_->tableWidget_unverifiableTests->removeRow(this->ui_->tableWidget_unverifiableTests->currentRow());
}

void TestsDialog::resizeEvent(QResizeEvent *event)
{
    this->ui_->tableWidget_verifiableTests->resize(this->width() - 20,
                                                   this->height() / 2 - this->ui_->tableWidget_verifiableTests->y() - 5);
    this->ui_->line->move(0, this->height() / 2);
    this->ui_->line->setFixedWidth(this->width());

    this->ui_->label_unverifFilter->move(10, this->height() / 2 + 25);
    this->ui_->comboBox_unverifiableTests_filter->move(10, this->height() / 2 + 45);
    this->ui_->label_noControlParams->move(1400, this->height() / 2 + 20);
    this->ui_->comboBox_unverifiableTests->move(130, this->height() / 2 + 45);
    this->ui_->pushButton_appendUnVerifiableTest->move(580, this->height() / 2 + 30);
    this->ui_->pushButton_removeUnVerifiableTestFromTable->move(650, this->height() / 2 + 20);

    this->ui_->tableWidget_unverifiableTests->move(10, this->height() / 2 + 75);
    this->ui_->tableWidget_unverifiableTests->resize(this->width() - 20, this->height() / 2 - 95);
}

void TestsDialog::on_comboBox_verifiableTests_filter_currentIndexChanged(const QString &arg1)
{
    this->ui_->comboBox_verifiableTests->clear();
    if (arg1 == "Все тесты")
        this->ui_->comboBox_verifiableTests->addItems(this->slTests_);
    else
    {
        QString filter = arg1.toLower();
        while ((arg1.size() / 2 <= filter.size()) && (filter.size() > 3))
            filter.remove(filter.size() - 1, 1);
        if (!this->slTests_.isEmpty())
            for (int i = 0; i < this->slTests_.size(); i++)
            {
                if (this->slTests_.at(i).toLower().contains(filter))
                    this->ui_->comboBox_verifiableTests->addItem(this->slTests_.at(i));
            }
    }
}

void TestsDialog::on_comboBox_unverifiableTests_filter_currentIndexChanged(const QString &arg1)
{
    this->ui_->comboBox_unverifiableTests->clear();
    if (arg1 == "Все тесты")
        this->ui_->comboBox_unverifiableTests->addItems(this->slTests_);
    else
    {
        QString filter = arg1.toLower();
        while ((arg1.size() / 2 <= filter.size()) && (filter.size() > 3))
            filter.remove(filter.size() - 1, 1);
        if (!this->slTests_.isEmpty())
            for (int i = 0; i < this->slTests_.size(); i++)
            {
                if (this->slTests_.at(i).toLower().contains(filter))
                    this->ui_->comboBox_unverifiableTests->addItem(this->slTests_.at(i));
            }
    }
}
