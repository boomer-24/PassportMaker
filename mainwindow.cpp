#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *_parent) : QMainWindow(_parent), ui_(new Ui::MainWindow)
{
    this->ui_->setupUi(this);
    this->setWindowTitle("Паспортогенератор");
    this->setFixedSize(630, 690);
    if (QFile::exists("animals&formats.xml"))
    {
        this->Initialize("animals&formats.xml");
        this->ui_->textEdit_commentKY_tt->setText(this->commentConnectionKYTT_);
        if (!this->allAnimals_.isEmpty())
        {
            this->ui_->comboBox_devName_2k->addItems(this->allAnimals_);
            this->ui_->comboBox_devName_tt->addItems(this->allAnimals_);
        }
        if (!this->animalsPositions_.isEmpty())
        {
            this->ui_->comboBox_devPosition_2k->addItems(this->animalsPositions_);
            this->ui_->comboBox_devPosition_tt->addItems(this->animalsPositions_);
        }
        this->ui_->lineEdit_inputAdapterName_2k->setText(this->adapterName2K_);
        this->ui_->lineEdit_inputAdapterNumber_2k->setText(this->adapterNumber2K_);
    } else this->ui_->textBrowser_info_2k->append("DAMN!!! Нет \"animals&formats.xml\"  в корне!");
    if (QFile::exists("tests&filters.xml"))
    {
        this->InitializeTestsTT("tests&filters.xml");
    } else this->ui_->textBrowser_info_2k->append("DAMN!!! Нет \"tests&filters.xml\"  в корне!");

    this->ui_->pushButton_openDirTT->setHidden(true);
    this->ui_->pushButton_openDir2K->setHidden(true);

    this->ui_->pushButton_jumpers->setEnabled(true);
    this->ui_->radioButton_jumpersYes->setChecked(true);

    QObject::connect(&this->lastCheckDialog2K_, SIGNAL(signalCreatePassport()), this, SLOT(slotCreatePassport2K()));
    QObject::connect(&this->lastCheckDialogTT_, SIGNAL(signalCreatePassport()), this, SLOT(slotCreatePassportTT()));
}

MainWindow::~MainWindow()
{
    this->SaveTestXml(this->testDialog_.GetTests(), this->testDialog_.GetFilters());
//    this->threadTT_.quit(); // ТОБЫ БЫЛА МНОГОПОТОЧНОСТЬ С WINWORD НУЖНО НАПИСАТЬ CoEx где-то
    delete ui_;
}

bool MainWindow::isOnlySpace(const QString &_str)
{
    for (auto& character : _str)
        if (!character.isSpace() && !character.isNonCharacter() && !character.isPunct() && !character.isDigit())
            return false;
    return true;
}

void MainWindow::Initialize(const QString &_path)
{
    QDomDocument domDoc;
    QFile file(_path);
    if (file.open(QIODevice::ReadOnly))
    {
        if (domDoc.setContent(&file))
        {
            QDomElement domElement = domDoc.documentElement();
            QDomNode domNode = domElement.firstChild();

            while(!domNode.isNull())
            {
                if(domNode.isElement())
                {
                    QDomElement domElement = domNode.toElement();
                    if(!domElement.isNull())
                    {
                        if(domElement.tagName() == "analogChief") this->analogChief_ = domElement.text();
                        else if(domElement.tagName() == "digitChief") this->digitChief_ = domElement.text();
                        else if(domElement.tagName() == "analogAnimal")
                        {
                            this->analogAnimals_.push_back(domElement.text());
                            this->allAnimals_.push_back(domElement.text());
                        }
                        else if(domElement.tagName() == "digitAnimal")
                        {
                            this->digitAnimals_.push_back(domElement.text());
                            this->allAnimals_.push_back(domElement.text());
                        }
                        else if(domElement.tagName() == "animalPosition") this->animalsPositions_.push_back(domElement.text());
                        else if(domElement.tagName() == "defaultFormat_2K")
                        {
                            this->slFormats2K_.push_back(domElement.text());
                            this->ui_->listWidget_enteredFormats_2k->addItem(domElement.text());
                            this->ui_->listWidget_enteredFormats_2k->setFixedHeight(35 + (this->slFormats2K_.size() - 1) * 25);
                        }
                        else if(domElement.tagName() == "tester_2K")
                        {
                            this->connectionDialog_.InsertTesterFromXML(domElement.text());
                            this->slTesters2K_.push_back(domElement.text());
                        }
                        else if(domElement.tagName() == "defaultFormat_TT")
                        {
                            this->slFormatsTT_.push_back(domElement.text());
                            this->ui_->listWidget_enteredFormats_tt->addItem(domElement.text());
                        }
                        else if(domElement.tagName() == "adapterName_2K") this->adapterName2K_ = domElement.text();
                        else if(domElement.tagName() == "adapterNumber_2K") this->adapterNumber2K_ = domElement.text();
                        else if(domElement.tagName() == "commentKY_TT") this->commentConnectionKYTT_ = domElement.text();
                        else qDebug() << "Tag ne find";
                    }
                    domNode = domNode.nextSibling();
                }
            }
        } else qDebug() << "It`s no XML!";
    } else qDebug() << "File not open =/";
}

bool MainWindow::isWordRunning()
{
    QString AppPath(QCoreApplication::applicationDirPath());
    AppPath.append("/tasklist.txt");
    QString query("tasklist > ");
    query.append(AppPath);
    system(query.toStdString().data());
    QFile file(AppPath);
    if (file.open(QIODevice::ReadOnly))
    {
        QByteArray data;
        data = file.readAll();
        QString str(data);
        str = str.toLower();
        if (str.contains("word"))
            return true;
        else
            return false;
    }
}

void MainWindow::InitializeTestsTT(const QString &_path)
{
    QDomDocument domDoc;
    QFile file(_path);
    if (file.open(QIODevice::ReadOnly))
    {
        if (domDoc.setContent(&file))
        {
            QDomElement domElement = domDoc.documentElement();
            QDomNode domNode = domElement.firstChild();
            QStringList tests;
            QStringList filters;
            while(!domNode.isNull())
            {
                if(domNode.isElement())
                {
                    QDomElement domElement = domNode.toElement();
                    if(!domElement.isNull())
                    {
                        if(domElement.tagName() == "test_TT") tests.push_back(domElement.text());
                        if(domElement.tagName() == "filter_TT") filters.push_back(domElement.text());
                        else qDebug() << "Tag ne find";
                    }
                    domNode = domNode.nextSibling();
                }
            }
            this->testDialog_.SetTests(tests);
            this->testDialog_.SetFilters(filters);
        } else qDebug() << "It`s no XML!";
    } else qDebug() << "File not open =/";
}

QDomElement MainWindow::MakeElement(QDomDocument &_doc, const QString &_name, const QString &_value)
{
    QDomElement domElement = _doc.createElement(_name);
    if (!_value.isEmpty())
    {
        QDomText domText = _doc.createTextNode(_value);
        domElement.appendChild(domText);
    }
    return domElement;
}

void MainWindow::SaveTestXml(const QStringList &_slTests, const QStringList &_slFilters)
{
    QFile file("tests&filters.xml");
    if (file.open(QIODevice::WriteOnly))
    {
        QXmlStreamWriter xmlWriter(&file);
        xmlWriter.setAutoFormatting(true);
        xmlWriter.writeStartDocument();
        xmlWriter.writeStartElement("general");
        for (const QString& filter : _slFilters)
        {
            xmlWriter.writeStartElement("filter_TT");
            xmlWriter.writeCharacters(filter);
            xmlWriter.writeEndElement();
        }
        for (const QString& test : _slTests)
        {
            xmlWriter.writeStartElement("test_TT");
            xmlWriter.writeCharacters(test);
            xmlWriter.writeEndElement();
        }
        xmlWriter.writeEndElement();
        xmlWriter.writeEndDocument();
        file.close();
    }
}

void MainWindow::on_comboBox_devName_2k_activated(const QString &arg1)
{
    this->ui_->comboBox_devName_tt->setCurrentText(arg1);
    if (arg1.contains("Клык"))
        this->setCursor(QCursor(QPixmap("cursors/iliaCursor.png")));
    else if (arg1.contains("Шанг"))
        this->setCursor(QCursor(QPixmap("cursors/gryshaCursor.png")));
    else if (arg1.contains("Харит"))
        this->setCursor(QCursor(QPixmap("cursors/lordCursor.png")));
    else if (arg1.contains("Клю"))
        this->setCursor(QCursor(QPixmap("cursors/kotCursor.png")));
    else if (arg1.contains("Ковеш"))
        this->setCursor(QCursor(QPixmap("cursors/serkCursor.png")));
    else if (arg1.contains("дор"))
        this->setCursor(QCursor(QPixmap("cursors/dimkaCursor.png")));
    else if (arg1.contains("Кагир"))
        this->setCursor(QCursor(QPixmap("cursors/ksyuCursor.png")));
    else if (arg1.contains("Молот"))
        this->setCursor(QCursor(QPixmap("cursors/molotCursor.png")));
    else if (arg1.isEmpty())
        this->setCursor(QCursor(QPixmap("cursors/molotCursor.png")));
}

void MainWindow::on_comboBox_devPosition_2k_activated(const QString &arg1)
{
    this->ui_->comboBox_devPosition_tt->setCurrentText(arg1);
}

void MainWindow::on_comboBox_devName_tt_activated(const QString &arg1)
{
    this->ui_->comboBox_devName_2k->setCurrentText(arg1);
    if (arg1.contains("Клык"))
        this->setCursor(QCursor(QPixmap("cursors/iliaCursor.png")));
    else if (arg1.contains("Шанг"))
        this->setCursor(QCursor(QPixmap("cursors/gryshaCursor.png")));
    else if (arg1.contains("Харит"))
        this->setCursor(QCursor(QPixmap("cursors/lordCursor.png")));
    else if (arg1.contains("Клю"))
        this->setCursor(QCursor(QPixmap("cursors/kotCursor.png")));
    else if (arg1.contains("Ковеш"))
        this->setCursor(QCursor(QPixmap("cursors/serkCursor.png")));
    else if (arg1.contains("дор"))
        this->setCursor(QCursor(QPixmap("cursors/dimkaCursor.png")));
    else if (arg1.contains("Кагир"))
        this->setCursor(QCursor(QPixmap("cursors/ksyuCursor.png")));
    else if (arg1.contains("Молот"))
        this->setCursor(QCursor(QPixmap("cursors/molotCursor.png")));
    else if (arg1.isEmpty())
        this->setCursor(QCursor(QPixmap("cursors/molotCursor.png")));

}

void MainWindow::on_comboBox_devPosition_tt_activated(const QString &arg1)
{
    this->ui_->comboBox_devPosition_2k->setCurrentText(arg1);
}

void MainWindow::on_pushButton_addElement_2k_clicked()
{
    QString element = this->ui_->lineEdit_inputElement_2k->text().toUpper();
    if (!this->isOnlySpace(element))
    {
        if (!this->slElements2K_.contains(element) && !element.isEmpty())
        {
            this->slElements2K_.push_back(element);
            this->ui_->listWidget_enteredElements_2k->addItem(element);
            this->ui_->label_elements_2k->setText(QString("Элементов ").append(QString::number(this->slElements2K_.size())));
            if (this->slElements2K_.size() > 1)
            {
                this->ui_->listWidget_enteredElements_2k->setFixedHeight(30 + (this->slElements2K_.size() - 1) * 25);
                if (this->ui_->listWidget_enteredElements_2k->height() > 120)
                    this->ui_->listWidget_enteredElements_2k->setFixedHeight(120);
            }
        }
    }
    this->ui_->lineEdit_inputElement_2k->clear();
}

void MainWindow::on_pushButton_removeElement_2k_clicked()
{
    if (this->slElements2K_.size() > 0 && this->ui_->listWidget_enteredElements_2k->currentItem() != nullptr)
    {
        QString element = this->ui_->listWidget_enteredElements_2k->currentItem()->text();
        this->slElements2K_.removeOne(element);

        this->ui_->listWidget_enteredElements_2k->removeItemWidget(this->ui_->listWidget_enteredElements_2k->takeItem(this->ui_->listWidget_enteredElements_2k->currentRow()));
        this->ui_->label_elements_2k->setText(QString("Элементов ").append(QString::number(this->slElements2K_.size())));

        this->ui_->listWidget_enteredElements_2k->setFixedHeight(30 + (this->slElements2K_.size() - 1) * 25);
        if (this->ui_->listWidget_enteredElements_2k->height() > 120)
            this->ui_->listWidget_enteredElements_2k->setFixedHeight(120);
        if (this->slElements2K_.size() == 0)
            this->ui_->listWidget_enteredElements_2k->setFixedHeight(30);
    }
}

void MainWindow::on_pushButton_addFormat_2k_clicked()
{
    if (!this->ui_->lineEdit_inputFormat_2k->text().isEmpty())
    {
        QString nakedFormat = this->ui_->lineEdit_inputFormat_2k->text().toLower();
        if (!this->isOnlySpace(nakedFormat))
        {
            if (!this->slFormats2K_.contains(nakedFormat) && !nakedFormat.isEmpty())
            {
                this->slFormats2K_.push_back(nakedFormat);
                this->ui_->listWidget_enteredFormats_2k ->addItem(nakedFormat);
                this->ui_->label_format_2k->setText(QString("Форматов ").append(QString::number(this->slFormats2K_.size())));
            }
            if (this->slFormats2K_.size() > 1)
            {
                this->ui_->listWidget_enteredFormats_2k->setFixedHeight(30 + (this->slFormats2K_.size() - 1) * 25);
                if (this->ui_->listWidget_enteredFormats_2k->height() > 120)
                    this->ui_->listWidget_enteredFormats_2k->setFixedHeight(120);
            }
        }
        this->ui_->lineEdit_inputFormat_2k->clear();
    }
}

void MainWindow::on_pushButton_removeFormat_2k_clicked()
{
    if (this->slFormats2K_.size() > 0 && this->ui_->listWidget_enteredFormats_2k->currentItem() != nullptr)
    {
        QString format = this->ui_->listWidget_enteredFormats_2k->currentItem()->text();
        this->slFormats2K_.removeOne(format);

        this->ui_->listWidget_enteredFormats_2k->removeItemWidget(this->ui_->listWidget_enteredFormats_2k->takeItem(this->ui_->listWidget_enteredFormats_2k->currentRow()));

        this->ui_->listWidget_enteredFormats_2k->setFixedHeight(30 + (this->slFormats2K_.size() - 1) * 25);
        if (this->ui_->listWidget_enteredFormats_2k->height() > 120)
            this->ui_->listWidget_enteredFormats_2k->setFixedHeight(120);
        if (this->slFormats2K_.size() == 0)
            this->ui_->listWidget_enteredFormats_2k->setFixedHeight(30);
        this->ui_->label_format_2k->setText(QString("Форматов ").append(QString::number(this->slFormats2K_.size())));
    }
}

void MainWindow::on_pushButton_selectDirectory_2k_clicked()
{
    this->dirPath_2K_ = (QFileDialog::getExistingDirectory(0, "Select directory"));
    this->ui_->lineEdit_dirPath_2k->setText(this->dirPath_2K_);
}

void MainWindow::on_pushButton_createSavePasport_2k_clicked()
{
    if (!this->isWordRunning())
    {
        if (this->ui_->lineEdit_dirPath_2k->text().isNull() || this->ui_->lineEdit_dirPath_2k->text() == " " ||
                this->ui_->lineEdit_dirPath_2k->text() == "")
        {
            QMessageBox::StandardButton reply = QMessageBox::warning(this, "Warning", "Не указана папка с программой!\n"
                                                                                      "Без этого никак.");
            return;
        }
        QVector<QPair<QString, QString>> vControlMessage;

        (this->ui_->comboBox_devName_2k->currentText() == " " || this->ui_->comboBox_devName_2k->currentText().isNull() ||
                this->ui_->comboBox_devName_2k->currentText() == "") ?
                    vControlMessage.push_back(QPair<QString, QString>("Имя разработчика", "")) :
                    vControlMessage.push_back(QPair<QString, QString>("Имя разработчика", this->ui_->comboBox_devName_2k->currentText()));

        (this->ui_->comboBox_devPosition_2k->currentText() == " " || this->ui_->comboBox_devPosition_2k->currentText().isNull() ||
                this->ui_->comboBox_devPosition_2k->currentText() == "") ?
                    vControlMessage.push_back(QPair<QString, QString>("Должность разработчика", "")) :
                    vControlMessage.push_back(QPair<QString, QString>("Должность разработчика", this->ui_->comboBox_devPosition_2k->currentText()));

        (this->ui_->lineEdit_inputTY_2k->text() == "") ?
                    vControlMessage.push_back(QPair<QString, QString>("ТУ", "")) :
                    vControlMessage.push_back(QPair<QString, QString>("ТУ", this->ui_->lineEdit_inputTY_2k->text()));

        (this->ui_->lineEdit_inputCorrectionTY_2k->text() == "") ?
                    vControlMessage.push_back(QPair<QString, QString>("Коррекция ТУ", "")) :
                    vControlMessage.push_back(QPair<QString, QString>("Коррекция ТУ", this->ui_->lineEdit_inputCorrectionTY_2k->text()));

        (this->ui_->lineEdit_inputAdapterName_2k->text() == "") ?
                    vControlMessage.push_back(QPair<QString, QString>("Адаптер", "")) :
                    vControlMessage.push_back(QPair<QString, QString>("Адаптер", this->ui_->lineEdit_inputAdapterName_2k->text()));

        (this->ui_->lineEdit_inputAdapterNumber_2k->text() == "") ?
                    vControlMessage.push_back(QPair<QString, QString>("Номер адаптера", "")) :
                    vControlMessage.push_back(QPair<QString, QString>("Номер адаптера", this->ui_->lineEdit_inputAdapterNumber_2k->text()));

        (this->ui_->lineEdit_inputKYName_2k->text() == "") ?
                    vControlMessage.push_back(QPair<QString, QString>("КУ", "")) :
                    vControlMessage.push_back(QPair<QString, QString>("КУ", this->ui_->lineEdit_inputKYName_2k->text()));

        (this->ui_->lineEdit_inputKYNumber_2k->text() == "") ?
                    vControlMessage.push_back(QPair<QString, QString>("Номер КУ", "")) :
                    vControlMessage.push_back(QPair<QString, QString>("Номер КУ", this->ui_->lineEdit_inputKYNumber_2k->text()));

        if (this->slElements2K_.empty())
            vControlMessage.push_back(QPair<QString, QString>("Проверяемые элементы", ""));
        else
        {
            QString elements;
            for (int i = 0; i < this->slElements2K_.size(); i++)
            {
                const QString& element = this->slElements2K_.at(i);
                if (0 < i && i < this->slElements2K_.size())
                    elements.append(" // ");
                elements.append(element);
            }
            vControlMessage.push_back(QPair<QString, QString>("Проверяемые элементы", elements));
        }

        if (this->ui_->radioButton_jumpersYes->isChecked())
        {
            if (!this->jumpersDialog_.GetTableContent().isEmpty())
            {
                QString jumpersValue;
                QVector<QPair<QString, QString>> rvTableContent = this->jumpersDialog_.GetTableContent();
                for (int i = 0; i < rvTableContent.size(); i++)
                {
                    if (0 < i)
                        jumpersValue.append(" || ");
                    const QPair<QString, QString>& pair = rvTableContent.at(i);
                    jumpersValue.append(pair.first).append(" : ").append(pair.second);
                }
                vControlMessage.push_back(QPair<QString, QString>("Перемычки", jumpersValue));
            } else vControlMessage.push_back(QPair<QString, QString>("Перемычек", "Нет"));
        } else vControlMessage.push_back(QPair<QString, QString>("Перемычек", "Нет"));

        this->lastCheckDialog2K_.FillTable(vControlMessage);
        this->lastCheckDialog2K_.show();
    } else
        QMessageBox::StandardButton reply = QMessageBox::warning(this, "Warning",
                                                                 "Процессы будут конфликтовать.\nДля дальнейшей работы закрой програму Word!\n"
                                                                 "(Если приложение MS Word не запущено, а ты видишь меня, "
                                                                 "значит у тебя включена возможность предварительного провмотра в папках, "
                                                                 "которая запускает процесс WINWORD.\n"
                                                                 "В этом случае необходимо завершить процесс WINWORD.EXE вручную "
                                                                 "в диспетчере задач)");

}

void MainWindow::on_pushButton_addElement_tt_clicked()
{
    QString element = this->ui_->lineEdit_inputElement_tt->text().toUpper();
    if (!this->isOnlySpace(element))
    {
        if (!this->slElementsTT_.contains(element) && !element.isEmpty())
        {
            this->slElementsTT_.push_back(element);
            this->ui_->listWidget_enteredElements_tt->addItem(element);
            this->ui_->label_elements_tt->setText(QString("Элементов ").append(QString::number(this->slElementsTT_.size())));
            if (this->slElementsTT_.size() > 1)
            {
                this->ui_->listWidget_enteredElements_tt->setFixedHeight(30 + (this->slElementsTT_.size() - 1) * 25);
                if (this->ui_->listWidget_enteredElements_tt->height() > 120)
                    this->ui_->listWidget_enteredElements_tt->setFixedHeight(120);
            }
        }
    }
    this->ui_->lineEdit_inputElement_tt->clear();
}

void MainWindow::on_pushButton_removeElement_tt_clicked()
{
    if (this->slElementsTT_.size() > 0 && this->ui_->listWidget_enteredElements_tt->currentItem() != nullptr)
    {
        QString element = this->ui_->listWidget_enteredElements_tt->currentItem()->text();
        this->slElementsTT_.removeOne(element);

        this->ui_->listWidget_enteredElements_tt->removeItemWidget(this->ui_->listWidget_enteredElements_tt->takeItem(this->ui_->listWidget_enteredElements_tt->currentRow()));
        this->ui_->label_elements_tt->setText(QString("Элементов ").append(QString::number(this->slElementsTT_.size())));

        this->ui_->listWidget_enteredElements_tt->setFixedHeight(30 + (this->slElementsTT_.size() - 1) * 25);
        if (this->ui_->listWidget_enteredElements_tt->height() > 120)
            this->ui_->listWidget_enteredElements_tt->setFixedHeight(120);
        if (this->slElementsTT_.size() == 0)
            this->ui_->listWidget_enteredElements_tt->setFixedHeight(30);
    }
}

void MainWindow::on_pushButton_addFormat_tt_clicked()
{
    if (!this->ui_->lineEdit_inputFormat_tt->text().isEmpty())
    {
        QString nakedFormat = this->ui_->lineEdit_inputFormat_tt->text().toLower();
        if (!this->isOnlySpace(nakedFormat))
        {
            if (!this->slFormatsTT_.contains(nakedFormat) && !nakedFormat.isEmpty())
            {
                this->slFormatsTT_.push_back(nakedFormat);
                this->ui_->listWidget_enteredFormats_tt ->addItem(nakedFormat);
                this->ui_->label_format_tt->setText(QString("Форматов ").append(QString::number(this->slFormatsTT_.size())));
            }
            if (this->slFormatsTT_.size() > 1)
            {
                this->ui_->listWidget_enteredFormats_tt->setFixedHeight(30 + (this->slFormatsTT_.size() - 1) * 25);
                if (this->ui_->listWidget_enteredFormats_tt->height() > 120)
                    this->ui_->listWidget_enteredFormats_tt->setFixedHeight(120);
            }
        }
        this->ui_->lineEdit_inputFormat_tt->clear();
    }
}

void MainWindow::on_pushButton_removeFormat_tt_clicked()
{
    if (this->slFormatsTT_.size() > 0 && this->ui_->listWidget_enteredFormats_tt->currentItem() != nullptr)
    {
        QString format = this->ui_->listWidget_enteredFormats_tt->currentItem()->text();
        this->slFormatsTT_.removeOne(format);

        this->ui_->listWidget_enteredFormats_tt->removeItemWidget(this->ui_->listWidget_enteredFormats_tt->takeItem(this->ui_->listWidget_enteredFormats_tt->currentRow()));

        this->ui_->listWidget_enteredFormats_tt->setFixedHeight(30 + (this->slFormatsTT_.size() - 1) * 25);
        if (this->ui_->listWidget_enteredFormats_tt->height() > 120)
            this->ui_->listWidget_enteredFormats_tt->setFixedHeight(120);
        if (this->slFormatsTT_.size() == 0)
            this->ui_->listWidget_enteredFormats_tt->setFixedHeight(30);
        this->ui_->label_format_tt->setText(QString("Форматов ").append(QString::number(this->slFormatsTT_.size())));
    }
}

void MainWindow::on_pushButton_selectDirectory_tt_clicked()
{
    this->dirPath_TT_ = (QFileDialog::getExistingDirectory(0, "Select directory"));
    this->ui_->lineEdit_dirPath_tt->setText(this->dirPath_TT_);
}

void MainWindow::on_pushButton_createSavePasport_tt_clicked()
{
    if (!this->isWordRunning())
    {
        if (this->ui_->lineEdit_dirPath_tt->text().isNull() || this->ui_->lineEdit_dirPath_tt->text() == " " ||
                this->ui_->lineEdit_dirPath_tt->text() == "")
        {
            QMessageBox::StandardButton reply = QMessageBox::warning(this, "Warning", "Не указана папка с программой!\n"
                                                                                      "Без этого никак.");
            return;
        }
        QVector<QPair<QString, QString>> vControlMessage;

        (this->ui_->comboBox_devName_tt->currentText() == " " || this->ui_->comboBox_devName_tt->currentText().isNull() ||
                this->ui_->comboBox_devName_tt->currentText() == "") ?
                    vControlMessage.push_back(QPair<QString, QString>("Имя разработчика", "")) :
                    vControlMessage.push_back(QPair<QString, QString>("Имя разработчика", this->ui_->comboBox_devName_tt->currentText()));

        (this->ui_->comboBox_devPosition_tt->currentText() == " " || this->ui_->comboBox_devPosition_tt->currentText().isNull() ||
                this->ui_->comboBox_devPosition_tt->currentText() == "") ?
                    vControlMessage.push_back(QPair<QString, QString>("Должность разработчика", "")) :
                    vControlMessage.push_back(QPair<QString, QString>("Должность разработчика", this->ui_->comboBox_devPosition_tt->currentText()));

        (this->ui_->lineEdit_inputTY_tt->text() == "") ?
                    vControlMessage.push_back(QPair<QString, QString>("ТУ", "")) :
                    vControlMessage.push_back(QPair<QString, QString>("ТУ", this->ui_->lineEdit_inputTY_tt->text()));

        (this->ui_->lineEdit_inputCorrectionTY_tt->text() == "") ?
                    vControlMessage.push_back(QPair<QString, QString>("Коррекция ТУ", "")) :
                    vControlMessage.push_back(QPair<QString, QString>("Коррекция ТУ", this->ui_->lineEdit_inputCorrectionTY_tt->text()));

        (this->ui_->lineEdit_inputKYNumber_tt->text() == "") ?
                    vControlMessage.push_back(QPair<QString, QString>("Номер КУ", "")) :
                    vControlMessage.push_back(QPair<QString, QString>("Номер КУ", this->ui_->lineEdit_inputKYNumber_tt->text()));

        (this->ui_->textEdit_commentKY_tt->toPlainText() == "") ?
                    vControlMessage.push_back(QPair<QString, QString>("Комментарий к подключению КУ", "")) :
                    (this->ui_->textEdit_commentKY_tt->toPlainText().contains("\n")) ?
                        vControlMessage.push_back(QPair<QString, QString>("Комментарий к подключению КУ",
                                                                          "Что-то написано, больше одной строки")) :
                        vControlMessage.push_back(QPair<QString, QString>("Комментарий к подключению КУ",
                                                                          this->ui_->textEdit_commentKY_tt->toPlainText()));

        (QFile::exists(this->ui_->lineEdit_dirPath_tt->text().append("/Безымянный.png"))) ?
                    vControlMessage.push_back(QPair<QString, QString>("Изображение как подключать КУ",
                                                                      this->ui_->lineEdit_dirPath_tt->text().append("/Безымянный.png"))) :
                    vControlMessage.push_back(QPair<QString, QString>("Изображение как подключать КУ", ""));

        this->lastCheckDialogTT_.FillTable(vControlMessage);
        this->lastCheckDialogTT_.show();
    } else
        QMessageBox::StandardButton reply = QMessageBox::warning(this, "Warning",
                                                                 "Процессы будут конфликтовать.\nДля дальнейшей работы закрой програму Word!\n"
                                                                 "(Если приложение MS Word не запущено, а ты видишь меня, "
                                                                 "значит у тебя включена возможность предварительного провмотра в папках, "
                                                                 "которая запускает процесс WINWORD.\n"
                                                                 "В этом случае необходимо завершить процесс WINWORD.EXE вручную "
                                                                 "в диспетчере задач)");
}

void MainWindow::on_pushButton_openPaint_tt_clicked()
{
    QProcess process(this);
    process.startDetached("paint.bat");//либо в деструкторе открывать
}

void MainWindow::resizeEvent(QResizeEvent *event)
{
    this->ui_->tabWidget->setFixedSize(this->size() - QSize(10, 10));
}

void MainWindow::slotCreatePassport2K()
{
    PassportMaker2K Passport2K;
    Passport2K.Initialize("content.xml");
    if (this->analogAnimals_.contains(this->ui_->comboBox_devName_2k->currentText()))
        Passport2K.SetGroupChief(this->analogChief_);
    else
        if (this->digitAnimals_.contains(this->ui_->comboBox_devName_2k->currentText()))
            Passport2K.SetGroupChief(this->digitChief_);

    Passport2K.SetAnimal(this->ui_->comboBox_devName_2k->currentText());
    Passport2K.SetAnimalPosition(this->ui_->comboBox_devPosition_2k->currentText());
    Passport2K.SetTY(this->ui_->lineEdit_inputTY_2k->text());
    Passport2K.SetCorrectionTY(this->ui_->lineEdit_inputCorrectionTY_2k->text());
    Passport2K.SetAdapterName(this->ui_->lineEdit_inputAdapterName_2k->text());
    Passport2K.SetAdapterNumber(this->ui_->lineEdit_inputAdapterNumber_2k->text());
    Passport2K.SetKYname(this->ui_->lineEdit_inputKYName_2k->text());
    Passport2K.SetKYnumber(this->ui_->lineEdit_inputKYNumber_2k->text());

    Passport2K.setStrAndVvTableConnection(this->connectionDialog_.getTableContent());

    Passport2K.InitializeElementsSl(this->slElements2K_);
    Passport2K.InitializeFileFormatSl(this->slFormats2K_);

    (this->ui_->radioButton_control2k_2k->isChecked()) ? Passport2K.SetMode("Control2K") : (Passport2K.SetMode("Memory"));
    (this->ui_->radioButton_pinH_2k->isChecked()) ? Passport2K.SetPinH(true) : Passport2K.SetPinH(false);

    Passport2K.SetDirPath(this->ui_->lineEdit_dirPath_2k->text());

    if (this->ui_->radioButton_jumpersYes->isChecked())
        Passport2K.SetVJPos(this->jumpersDialog_.GetTableContent());

    //    Pasport2K

    Passport2K.CreateDocument();
    Passport2K.SaveDocx();

    this->ui_->textBrowser_info_2k->append(QString("Паспорт сохранен в ").append(this->dirPath_2K_));
    this->ui_->pushButton_openDir2K->setHidden(false);
}

void MainWindow::slotCreatePassportTT()
{
    PassportMakerTT PassportTT(this);
    QObject::connect(&this->lastCheckDialogTT_, SIGNAL(signalCreatePassport()), &PassportTT, SLOT(slotCreateAndSavePassport()));

    PassportTT.Initialize("content.xml");

    if (this->analogAnimals_.contains(this->ui_->comboBox_devName_2k->currentText()))
        PassportTT.SetGroupChief(this->analogChief_);
    else
        if (this->digitAnimals_.contains(this->ui_->comboBox_devName_2k->currentText()))
            PassportTT.SetGroupChief(this->digitChief_);

    PassportTT.SetAnimal(this->ui_->comboBox_devName_tt->currentText());
    PassportTT.SetAnimalPosition(this->ui_->comboBox_devPosition_tt->currentText());
    PassportTT.SetTY(this->ui_->lineEdit_inputTY_tt->text());
    PassportTT.SetCorrectionTY(this->ui_->lineEdit_inputCorrectionTY_tt->text());
    PassportTT.SetKYnumber(this->ui_->lineEdit_inputKYNumber_tt->text());
    PassportTT.SetKYComment(this->ui_->textEdit_commentKY_tt->toPlainText());

    PassportTT.InitializeElementsSl(this->slElementsTT_);
    PassportTT.InitializeFileFormatSl(this->slFormatsTT_);

    PassportTT.SetDirPath(this->ui_->lineEdit_dirPath_tt->text());

    PassportTT.SetVVerifiableTests(this->testDialog_.GetVerifiableTests());
    PassportTT.SetVUnVerifiableTests(this->testDialog_.GetUnVerifiableTests());

//    PassportTT.moveToThread(&this->threadTT_);
//    this->threadTT_.start();// ТОБЫ БЫЛА МНОГОПОТОЧНОСТЬ С WINWORD НУЖНО НАПИСАТЬ CoEx где-то

    PassportTT.CreateDocument();
    PassportTT.SaveDocx();

    this->ui_->textBrowser_info_tt->append(QString("Паспорт сохранен в ").append(this->dirPath_TT_));
    this->ui_->pushButton_openDirTT->setHidden(false);
}

void MainWindow::on_pushButton_testDialogBox_clicked()
{
    this->testDialog_.show();
}

void MainWindow::on_pushButton_openDirTT_clicked()
{
    QDir Dir(this->dirPath_TT_);
    QDesktopServices::openUrl(QUrl::fromLocalFile(Dir.absolutePath()));
}

void MainWindow::on_pushButton_openDir2K_clicked()
{
    QDir Dir(this->dirPath_2K_);
    QDesktopServices::openUrl(QUrl::fromLocalFile(Dir.absolutePath()));
}

void MainWindow::on_pushButton_jumpers_clicked()
{
    this->jumpersDialog_.show();
}

void MainWindow::on_radioButton_jumpersYes_clicked()
{
    this->ui_->pushButton_jumpers->setEnabled(true);
}

void MainWindow::on_radioButton_jumpersNo_clicked()
{
    this->ui_->pushButton_jumpers->setEnabled(false);
}

void MainWindow::on_pushButton_connectionDialog_clicked()
{
    this->connectionDialog_.show();
}
