#include "passportmakertt.h"

PassportMakerTT::PassportMakerTT(QObject* _parent) : QObject(_parent)
{
    this->wordApplication_ = new QAxObject("Word.Application");
    //    this->marker_ = true;
    this->wordDocument_ = this->wordApplication_->querySubObject("Documents()");
    this->wordDocument_->querySubObject("Add()");
    this->activeDocument_ = this->wordApplication_->querySubObject("ActiveDocument()");
    this->cryptHash_ = new QCryptographicHash(QCryptographicHash::Sha1);

    this->selection_ = this->wordApplication_->querySubObject("Selection()");
    this->range_ = selection_->querySubObject("Range()");
    this->tables_ = this->activeDocument_->querySubObject("Tables()");
    this->table_ = tables_->querySubObject("Add(Range,NumRows,NumColumns, DefaulttableBehavior, AutoFitBehavior)", range_->asVariant(), 2, 2);
    this->paragraphFormat_ = this->selection_->querySubObject("ParagraphFormat()");
    this->font_ = selection_->querySubObject("Font");
}

PassportMakerTT::~PassportMakerTT()
{
    this->wordApplication_->dynamicCall("Quit()");
    delete this->wordApplication_;
}

QString PassportMakerTT::getImagePath() const
{
    return imagePath_;
}

void PassportMakerTT::slotCreateAndSavePassport()
{
    this->CreateDocument();
    this->SaveDocx();
}

void PassportMakerTT::Initialize(const QString &_pathToXML)
{
    QDomDocument domDoc;
    QFile file(_pathToXML);
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
                        if(domElement.tagName() == "approveTrif") this->approveTrif_ = domElement.text();
                        else if(domElement.tagName() == "approveShat") this->approveShat_ = domElement.text();
                        else if(domElement.tagName() == "approveGroupChief1") this->approveGroupChief_1_ = domElement.text();
                        else if(domElement.tagName() == "approveGroupChief2") this->approveGroupChief_2_ = domElement.text();
                        else if(domElement.tagName() == "title_TT") this->title_ = domElement.text();
                        else if(domElement.tagName() == "titlePasport_TT") this->titlePasport_ = domElement.text();
                        else if(domElement.tagName() == "titleOther_TT") this->titleOther_ = domElement.text();
                        else if(domElement.tagName() == "firstParagraph1stString_TT") this->firstParagraph1stString_ = domElement.text();
                        else if(domElement.tagName() == "firstParagraph2ndString_TT") this->firstParagraph2ndString_ = domElement.text();
                        else if(domElement.tagName() == "otherOnTheFirstList_TT") this->otherOnThe1stList_ = domElement.text();
                        else if(domElement.tagName() == "description_TT") this->description_ = domElement.text();
                        else if(domElement.tagName() == "description1_TT") this->description1_ = domElement.text();
                        else if(domElement.tagName() == "instruction_TT") this->instruction_ = domElement.text();
                        else if(domElement.tagName() == "instruction1_TT") this->instruction1_ = domElement.text();
                        else if(domElement.tagName() == "guide_TT") this->guide_ = domElement.text();
                        else if(domElement.tagName() == "guide1_TT") this->guide1_ = domElement.text();
                        else qDebug() << "Tag ne find";
                    }
                    domNode = domNode.nextSibling();
                }
            }
        } else qDebug() << "It`s no XML!";
    } else qDebug() << "File not open =/";
}

void PassportMakerTT::InitializeElementsSl(const QStringList &_slElements)
{
    this->slElements_ = _slElements;
}

void PassportMakerTT::InitializeFileFormatSl(const QStringList &_slFileFormats)
{
    this->slFileFormats_ = _slFileFormats;
}

void PassportMakerTT::CreateDocument()
{
    this->TitleListCreate();
    this->SecondListCreate();
    this->OneMoreListCreate();
    this->TestersList();
//    this->AddHashSheet();
}

void PassportMakerTT::SetGroupChief(const QString &_chiefName)
{
    this->selectedChief_ = _chiefName;
}

void PassportMakerTT::SetAnimal(const QString &_animalName)
{
    this->selectedAnimal_ = _animalName;
}

void PassportMakerTT::SetAnimalPosition(const QString &_animalPosition)
{
    this->selectedPosition_ = _animalPosition;
}

void PassportMakerTT::PushVerifiableTest(const Test &_test)
{
    this->vVerifiableTests_.push_back(_test);
}

void PassportMakerTT::PushUnVerifiableTest(const Test &_test)
{
    this->vUnVerifiableTests_.push_back(_test);
}

void PassportMakerTT::SetVVerifiableTests(const QVector<Test> &vVerifiableTests)
{
    this->vVerifiableTests_ = vVerifiableTests;
}

void PassportMakerTT::SetVUnVerifiableTests(const QVector<Test> &vUnVerifiableTests)
{
    this->vUnVerifiableTests_ = vUnVerifiableTests;
}

void PassportMakerTT::AddToElementsStringList(const QString &_element)
{
    this->slElements_.push_back(_element);
}

void PassportMakerTT::RemoveSpecificElement(const QString &_element)
{
    this->slElements_.removeOne(_element);
}

bool PassportMakerTT::ContainsElement(const QString &_element)
{
    if (this->slElements_.contains(_element)) return true;
    else return false;
}

void PassportMakerTT::AddToFormatsStringList(const QString &_format)
{
    this->slFileFormats_.push_back(_format.toLower());
}

void PassportMakerTT::RemoveSpecificFormat(const QString &_format)
{
    QString format = QString("*.").append(_format);
    this->slFileFormats_.removeOne(format);
}

bool PassportMakerTT::ContainsFormat(const QString &_format)
{
    if (this->slFileFormats_.contains(_format)) return true;
    else return false;
}

int PassportMakerTT::GetSlElementsSize() const
{
    int size = this->slElements_.size();
    return size;
}

int PassportMakerTT::GetSlFormatsSize() const
{
    int size = this->slFileFormats_.size();
    return size;
}

void PassportMakerTT::TitleListCreate()
{
    if (!this->selection_->isNull())
    {
        QAxObject* section = this->activeDocument_->querySubObject("Sections(Index)", 1);
        QAxObject* footersFP = section->querySubObject("Footers(WdHeaderFooterIndex)", "wdHeaderFooterFirstPage");
        QAxObject* rangeFoot = footersFP->querySubObject("Range()");
        rangeFoot->dynamicCall("InsertBefore(Text)", "");

        //        QAxObject* font = selection_->querySubObject("Font");
        font_->setProperty("Size", 14);    font_->setProperty("Name", "Times New Roman");
        selection_->dynamicCall("TypeText(const QString&)",this->approveTrif_);
        selection_->dynamicCall("moveRight()");

        font_->setProperty("Size", 14);    font_->setProperty("Name", "Times New Roman");
        selection_->dynamicCall("TypeText(const QString&)",this->approveShat_);
        selection_->dynamicCall("moveDown()");
        selection_->dynamicCall("moveLeft()");

        font_->setProperty("Size", 14);    font_->setProperty("Name", "Times New Roman");
        selection_->dynamicCall("TypeText(const QString&)",this->approveGroupChief_1_);
        selection_->dynamicCall("TypeText(const QString&)",this->selectedChief_.append("\n"));
        selection_->dynamicCall("TypeText(const QString&)",this->approveGroupChief_2_);

        selection_->dynamicCall("moveDown()");
        for (int C = 0; C < 3; C++)
            selection_->dynamicCall("TypeParagraph()");

        font_->setProperty("Size", 20);    font_->setProperty("Name", "Times New Roman");
        QAxObject* allign = selection_->querySubObject("ParagraphFormat()");
        allign->dynamicCall("Alignment", "wdAlignParagraphCenter");

        selection_->dynamicCall("TypeText(const QString&)",this->title_);
        selection_->dynamicCall("TypeParagraph()");
        font_->setProperty("Size", 18);
        selection_->dynamicCall("TypeText(const QString&)",this->titlePasport_);
        selection_->querySubObject("InlineShapes")->dynamicCall("AddHorizontalLineStandard()");
        font_->setProperty("Size", 20);
        selection_->dynamicCall("TypeText(const QString&)",this->titleOther_);
        selection_->dynamicCall("TypeParagraph()");
        font_->setProperty("Bold", true);
        for (QString str : this->slElements_)
        {
            selection_->dynamicCall("TypeText(const QString&)", str);
            selection_->dynamicCall("TypeParagraph()");
        }
        font_->setProperty("Bold", false);

        selection_->dynamicCall("InsertBreak()", "wdPageBreak"); //новый лист
        font_->setProperty("Size", 12);
    }else qDebug() << "damned selection";
}

void PassportMakerTT::SecondListCreate()
{
    this->paragraphFormat_->dynamicCall("Alignment", "wdAlignParagraphJustify");

    this->selection_->querySubObject("PageSetup()")->setProperty("LeftMargin", 75);    // так работает
    this->paragraphFormat_->setProperty("LeftIndent", 0);                             // и так =)
    this->paragraphFormat_->setProperty("FirstLineIndent", 0);                             // и так =)

    this->selection_->dynamicCall("TypeText(const QString&)", this->firstParagraph1stString_);

    QString element;
    for (int C = 0; C < this->slElements_.size(); C++)
    {
        element.append(this->slElements_.at(C));
        if (C < this->slElements_.size() - 1)
            element.append(", ");
    }

    QAxObject* section = this->activeDocument_->querySubObject("Sections(Index)", 1);
    QAxObject* footers = section->querySubObject("Footers(WdHeaderFooterIndex)", "wdHeaderFooterPrimary");
    QAxObject* rangeFoot = footers->querySubObject("Range()");
    rangeFoot->dynamicCall("InsertAfter(Text)", QString(element).append("                               Разработчик:    ").append(this->selectedAnimal_));

    QAxObject* pageNumbers = this->activeDocument_->querySubObject("Sections(Index)", 1)->querySubObject("Footers(WdHeaderFooterIndex)", "wdHeaderFooterPrimary")->querySubObject("PageNumbers()")->querySubObject("Add(PageNumberAlignment, FirstPage)", "wdAlignPageNumberCenter", false);

    font_->setProperty("Bold", true);
    selection_->dynamicCall("TypeText(const QString&)",element);
    font_->setProperty("Bold", false);
    selection_->dynamicCall("TypeText(const QString&)",this->firstParagraph2ndString_);
    font_->setProperty("Bold", true);
    selection_->dynamicCall("TypeText(const QString&)",this->TY_);
    selection_->dynamicCall("TypeText(const QString&)",", коррекция ");
    selection_->dynamicCall("TypeText(const QString&)",this->correctionTY_);
    font_->setProperty("Bold", false);
    selection_->dynamicCall("TypeText(const QString&)",this->otherOnThe1stList_);
    font_->setProperty("Bold", true);
    selection_->dynamicCall("TypeParagraph()");

    range_ = selection_->querySubObject("Range()");
    parForm_ = this->range_->querySubObject("ParagraphFormat()");
    parForm_->setProperty("SpaceAfter", 3);

    selection_->dynamicCall("TypeText(const QString&)",this->description_);
    selection_->dynamicCall("TypeParagraph()");
    font_->setProperty("Bold", false);

    selection_->dynamicCall("TypeText(const QString&)",this->description1_);
    font_->setProperty("Bold", true);
    selection_->dynamicCall("TypeParagraph()");
    //    selection2K_->dynamicCall("TypeParagraph()");

    selection_->dynamicCall("TypeText(const QString&)",this->instruction_);
    selection_->dynamicCall("TypeParagraph()");
    font_->setProperty("Bold", false);
    selection_->dynamicCall("TypeText(const QString&)",this->instruction1_);
    font_->setProperty("Bold", true);
    selection_->dynamicCall("TypeParagraph()");
    //    selection2K_->dynamicCall("TypeParagraph()");

    selection_->dynamicCall("TypeText(const QString&)",this->guide_);
    selection_->dynamicCall("TypeParagraph()");
    font_->setProperty("Bold", false);
    selection_->dynamicCall("TypeText(const QString&)",this->guide1_);

    selection_->dynamicCall("InsertBreak()", "wdPageBreak");
}

void PassportMakerTT::OneMoreListCreate()
{
    font_->setProperty("Bold", true);
    selection_->dynamicCall("TypeText(const QString&)", "Таблица 1. Перечень проверяемых параметров и отклонений от ТУ.");
    font_->setProperty("Bold", false);
    range_ = selection_->querySubObject("Range()");
    table_ = tables_->querySubObject("Add(Range,NumRows,NumColumns, DefaultTableBehavior, AutoFitBehavior)",
                                     range_->asVariant(), this->vVerifiableTests_.size() + 1, 5, 1, 2);

    QAxObject* cell_11 = table_->querySubObject("Cell(Row,  Column)", 1, 1);
    cell_11->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "4.4", "wdAdjustNone");
    QAxObject* rangeCell = cell_11->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Наименование параметра, единица измерения. Нумерация по ТУ");

    QAxObject* cell_12 = table_->querySubObject("Cell(Row,  Column)", 1, 2);
    cell_12->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "2.8", "wdAdjustNone");
    rangeCell = cell_12->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Обозначение");

    QAxObject* cell_13 = table_->querySubObject("Cell(Row,  Column)", 1, 3);
    cell_13->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "2.8", "wdAdjustNone");
    rangeCell = cell_13->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Норма");

    QAxObject* cell_14 = table_->querySubObject("Cell(Row,  Column)", 1, 4);
    cell_14->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "3.73", "wdAdjustNone");
    rangeCell = cell_14->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Условия");

    QAxObject* cell_15 = table_->querySubObject("Cell(Row,  Column)", 1, 5);
    cell_15->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "3.5", "wdAdjustNone");
    rangeCell = cell_15->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Отклонения от ТУ");

    for (int C = 0; C < this->vVerifiableTests_.size(); C++)
    {
        QAxObject* cell_21 = table_->querySubObject("Cell(Row,  Column)", 2 + C, 1);
        cell_21->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "4.4", "wdAdjustNone");
        QAxObject* cell_22 = table_->querySubObject("Cell(Row,  Column)", 2 + C, 2);
        cell_22->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "2.8", "wdAdjustNone");
        QAxObject* cell_23 = table_->querySubObject("Cell(Row,  Column)", 2 + C, 3);
        cell_23->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "2.8", "wdAdjustNone");
        QAxObject* cell_24 = table_->querySubObject("Cell(Row,  Column)", 2 + C, 4);
        cell_24->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "3.73", "wdAdjustNone");
        QAxObject* cell_25 = table_->querySubObject("Cell(Row,  Column)", 2 + C, 5);
        cell_25->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "3.5", "wdAdjustNone");
    }

    QAxObject* cella = table_->querySubObject("Cell(Row,  Column)", 1, 3);
    cella->dynamicCall("Split(NumRows, NumColumns)", 2, 1);

    for (int i = 0; i <= this->vVerifiableTests_.size(); i++)
    {
        QAxObject* cella_ = table_->querySubObject("Cell(Row,  Column)", i + 2, 3);
        cella_->dynamicCall("Split(NumRows, NumColumns)", 1, 2);
    }

    QAxObject* cell_23_ = table_->querySubObject("Cell(Row,  Column)", 2, 3);
    rangeCell = cell_23_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Не менее");
    QAxObject* cell_24_ = table_->querySubObject("Cell(Row,  Column)", 2, 4);
    rangeCell = cell_24_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Не более");

    for (int i = 0; i < this->vVerifiableTests_.size(); i++)
        for (int j = 0; j < 6; j++) //tables columns
        {
            QAxObject* cell = table_->querySubObject("Cell(Row,  Column)", i + 3, j + 1);
            rangeCell = cell->querySubObject("Range()");
            if (j == 4)
            {
                QString str = this->vVerifiableTests_.at(i).GetTest().at(j);
                str.replace("+-", "±");
                str.replace("CC", "°C");
                str.replace("СС", "°C");
                str.replace("\\", "\n");
                rangeCell->dynamicCall("InsertAfter(Text)", str);
            } else rangeCell->dynamicCall("InsertAfter(Text)", this->vVerifiableTests_.at(i).GetTest().at(j));
        }
    table_->setProperty("PreferredWidthType()", "wdPreferredWidthPercent");
    table_->setProperty("PreferredWidth", 100);

    for (int C = 0; C < this->vVerifiableTests_.size() + 100; C++)
        selection_->dynamicCall("moveDown()");
    selection_->dynamicCall("TypeParagraph()");

    font_->setProperty("Bold", true);
    selection_->dynamicCall("TypeText(const QString&)", "Таблица 2. Перечень непроверяемых параметров.");
    font_->setProperty("Bold", false);
    range_ = selection_->querySubObject("Range()");
    table_ = tables_->querySubObject("Add(Range,NumRows,NumColumns, DefaultTableBehavior, AutoFitBehavior)",
                                     range_->asVariant(), this->vUnVerifiableTests_.size() + 1, 4, 1, 2);

    for (int C = 0; C <= this->vUnVerifiableTests_.size(); C++)
    {
        QAxObject* ccell21 = table_->querySubObject("Cell(Row,  Column)", 1 + C, 1);
        ccell21->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "4.4", "wdAdjustNone");
        QAxObject* ccell22 = table_->querySubObject("Cell(Row,  Column)", 1 + C, 2);
        ccell22->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "2.8", "wdAdjustNone");
        QAxObject* ccell23 = table_->querySubObject("Cell(Row,  Column)", 1 + C, 3);
        ccell23->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "2.8", "wdAdjustNone");
        QAxObject* ccell24 = table_->querySubObject("Cell(Row,  Column)", 1 + C, 4);
        ccell24->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "7.23", "wdAdjustNone");
    }
    QAxObject* ccella = table_->querySubObject("Cell(Row,  Column)", 1, 3);
    ccella->dynamicCall("Split(NumRows, NumColumns)", 2, 1);

    for (int i = 0; i <= this->vUnVerifiableTests_.size(); i++)
    {
        QAxObject* ccella_ = table_->querySubObject("Cell(Row,  Column)", i + 2, 3);
        ccella_->dynamicCall("Split(NumRows, NumColumns)", 1, 2);
    }
    QAxObject* ccell11 = table_->querySubObject("Cell(Row,  Column)", 1, 1);
    ccell11->querySubObject("Range()")->dynamicCall("InsertAfter(Text)", "Наименование параметра, единица измерения.");

    QAxObject* ccell12 = table_->querySubObject("Cell(Row,  Column)", 1, 2);
    ccell12->querySubObject("Range()")->dynamicCall("InsertAfter(Text)", "Обозначение");

    QAxObject* ccell13 = table_->querySubObject("Cell(Row,  Column)", 1, 3);
    ccell13->querySubObject("Range()")->dynamicCall("InsertAfter(Text)", "Норма");

    QAxObject* ccell14 = table_->querySubObject("Cell(Row,  Column)", 1, 4);
    ccell14->querySubObject("Range()")->dynamicCall("InsertAfter(Text)", "Условия");

    QAxObject* ccell_23_ = table_->querySubObject("Cell(Row,  Column)", 2, 3);
    rangeCell = ccell_23_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Не менее");
    QAxObject* ccell_24_ = table_->querySubObject("Cell(Row,  Column)", 2, 4);
    rangeCell = ccell_24_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Не более");

    for (int i = 0; i < this->vUnVerifiableTests_.size(); i++)
        for (int j = 0; j < 5; j++) //tables columns
        {
            QAxObject* cell = table_->querySubObject("Cell(Row,  Column)", i + 3, j + 1);
            rangeCell = cell->querySubObject("Range()");
            if (j == 4)
            {
                QString str = this->vUnVerifiableTests_.at(i).GetTest().at(j);
                str.replace("+-", "±");
                str.replace("CC", "°C");
                str.replace("СС", "°C");
                str.replace("\\", "\n");
                rangeCell->dynamicCall("InsertAfter(Text)", str);
            } else rangeCell->dynamicCall("InsertAfter(Text)", this->vUnVerifiableTests_.at(i).GetTest().at(j));
        }

    for (int C = 0; C < this->vUnVerifiableTests_.size() + 100; C++)
        selection_->dynamicCall("moveDown()");
    selection_->dynamicCall("TypeParagraph()");

    font_->setProperty("Bold", true);
    selection_->dynamicCall("TypeText(const QString&)", "Таблица 3. Оснастка.");
    font_->setProperty("Bold", false);
    range_ = selection_->querySubObject("Range()");
    table_ = tables_->querySubObject("Add(Range,NumRows,NumColumns, DefaultTableBehavior, AutoFitBehavior)",
                                     range_->asVariant(), 2, 2, 1, 2);

    QAxObject* ccellDevice11 = table_->querySubObject("Cell(Row,  Column)", 1, 1);
    ccellDevice11->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "8.11", "wdAdjustNone");
    QAxObject* ccellDevice12 = table_->querySubObject("Cell(Row,  Column)", 1, 2);
    ccellDevice12->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "8.12", "wdAdjustNone");
    QAxObject* ccellDevice21 = table_->querySubObject("Cell(Row,  Column)", 2, 1);
    ccellDevice21->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "8.11", "wdAdjustNone");
    QAxObject* ccellDevice22 = table_->querySubObject("Cell(Row,  Column)", 2, 2);
    ccellDevice22->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "8.12", "wdAdjustNone");
    selection_->dynamicCall("TypeText(const QString&)","Тип оснастки");
    selection_->dynamicCall("moveRight()");
    selection_->dynamicCall("TypeText(const QString&)","Регистрационный номер оснастки");
    selection_->dynamicCall("moveDown()");
    selection_->dynamicCall("moveLeft()");
    selection_->dynamicCall("TypeText(const QString&)","ПКУ");
    selection_->dynamicCall("moveRight()");
    selection_->dynamicCall("TypeText(const QString&)",this->KYnumber_);

    selection_->dynamicCall("moveDown()");

    selection_->dynamicCall("TypeParagraph()");

    selection_->dynamicCall("InsertBreak()", "wdPageBreak"); //новый лист
}

void PassportMakerTT::TestersList()
{
    range_ = selection_->querySubObject("Range()");
    QAxObject* ParForm = this->range_->querySubObject("ParagraphFormat()");
    ParForm->setProperty("SpaceAfter", 0);

    font_->setProperty("Bold", true);
    selection_->dynamicCall("TypeText(const QString&)","Лист испытателя");   selection_->dynamicCall("TypeParagraph()");
    selection_->dynamicCall("TypeParagraph()");

    font_->setProperty("Bold", false);
    selection_->dynamicCall("TypeText(const QString&)", "Изделие: ");
    font_->setProperty("Bold", true);
    QString element;
    for (int i = 0; i < this->slElements_.size(); i++)
    {
        element.append(this->slElements_.at(i));
        if (i < this->slElements_.size() - 1) element.append(", ");
    }
    selection_->dynamicCall("TypeText(const QString&)", element);  selection_->dynamicCall("TypeParagraph()");

    font_->setProperty("Bold", false);
    selection_->dynamicCall("TypeText(const QString&)", "ТУ: ");
    font_->setProperty("Bold", true);
    selection_->dynamicCall("TypeText(const QString&)", this->TY_.append(", коррекция ").append(this->correctionTY_));   selection_->dynamicCall("TypeParagraph()");

    font_->setProperty("Bold", false);
    selection_->dynamicCall("TypeText(const QString&)", "Контактирующее устройство: ");
    font_->setProperty("Bold", true);
    selection_->dynamicCall("TypeText(const QString&)", this->KYnumber_);
    selection_->dynamicCall("TypeParagraph()");

    selection_->dynamicCall("TypeParagraph()");
    selection_->dynamicCall("TypeText(const QString&)", this->commentKYConnection_);
    selection_->dynamicCall("TypeParagraph()");
    font_->setProperty("Bold", false);

    this->selection_->dynamicCall("moveDown()");
    this->InsertImage();

    selection_->querySubObject("InlineShapes")->dynamicCall("AddHorizontalLineStandard()");

    this->AddHashSheet();
    this->selection_->dynamicCall("moveDown()");
    selection_->dynamicCall("TypeParagraph()");

    this->range_ = selection_->querySubObject("Range()");
    this->table_ = tables_->querySubObject("Add(Range,NumRows,NumColumns, DefaulttableBehavior, AutoFitBehavior)",
                                           this->range_->asVariant(), 2, 2);

    QAxObject* cell_11 = table_->querySubObject("Cell(Row,  Column)", 1, 1);
    cell_11->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "10.5", "wdAdjustNone");
    QAxObject* rangeCell = cell_11->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", this->selectedPosition_);
    QAxObject* cell_12 = table_->querySubObject("Cell(Row,  Column)", 1, 2);
    cell_12->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "5.5", "wdAdjustNone");
    rangeCell = cell_12->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", this->selectedAnimal_);

    QAxObject* cell_21 = table_->querySubObject("Cell(Row,  Column)", 2, 1);
    cell_21->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "10.5", "wdAdjustNone");
    QAxObject* cell_22 = table_->querySubObject("Cell(Row,  Column)", 2, 2);
    cell_22->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "5.5", "wdAdjustNone");
    rangeCell = cell_22->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "_______Дата");

    this->selection_->dynamicCall("moveDown()");
    selection_->dynamicCall("TypeParagraph()");
}

void PassportMakerTT::AddHashSheet()
{
    if (!this->slFiles_.isEmpty())
    {
        font_->setProperty("Size", 11);
        font_->setProperty("Name", "Calibri");

        font_->setProperty("Bold", true);
        selection_->dynamicCall("TypeText(const QString&)", "Таблица 6. Контрольные суммы файлов программы (SHA-1).");
        font_->setProperty("Bold", false);
        range_ = selection_->querySubObject("Range()");
        table_ = tables_->querySubObject("Add(Range, NumRows, NumColumns, DefaultTableBehavior, AutoFitBehavior)", range_->asVariant(), this->slFiles_.size() + 1, 3, 1, 0);

        QAxObject* cellTop0 = table_->querySubObject("Cell(Row, Column)", 1, 1);
        QAxObject* cellTopRange0 = cellTop0->querySubObject("Range()");
        cellTopRange0->dynamicCall("InsertAfter(Text)", "№ Версии");

        QAxObject* cellTop1 = table_->querySubObject("Cell(Row, Column)", 1, 2);
        QAxObject* cellTopRange1 = cellTop1->querySubObject("Range()");
        cellTopRange1->dynamicCall("InsertAfter(Text)", "Файл");

        QAxObject* cellTop2 = table_->querySubObject("Cell(Row, Column)", 1, 3);
        QAxObject* cellTopRange2 = cellTop2->querySubObject("Range()");
        cellTopRange2->dynamicCall("InsertAfter(Text)", "Контрольная сумма");

        for (int i = 0; i < this->slFiles_.size(); i++)
        {
            QAxObject* cell0 = table_->querySubObject("Cell(Row, Column)", i + 2, 1);
            QAxObject* cellRange0 = cell0->querySubObject("Range()");
            //            allign->dynamicCall("Alignment", "wdAlignParagraphCenter");
            cellRange0->dynamicCall("InsertAfter(Text)", "1");
            //            allign->dynamicCall("Alignment", "wdAlignParagraphLeft");

            QAxObject* cell1 = table_->querySubObject("Cell(Row, Column)", i + 2, 2);
            QAxObject* cellRange1 = cell1->querySubObject("Range()");
            QAxObject* cell2 = table_->querySubObject("Cell(Row, Column)", i + 2, 3);
            QAxObject* cellRange2 = cell2->querySubObject("Range()");
            cellRange1->dynamicCall("InsertAfter(Text)", this->slFiles_.at(i));

            QString path(this->directory_.path().append("/").append(this->slFiles_.at(i)));
            this->file_.setFileName(path);

            if (!this->file_.open(QIODevice::ReadOnly))
            {
                qDebug() << "file not open...";
            }else
            {
                this->cryptHash_->reset();
                this->data_ = this->file_.readAll();
                this->cryptHash_->addData(this->data_);
                this->result_ = this->cryptHash_->result().toHex().toUpper().data();
                cellRange2->dynamicCall("InsertAfter(Text)", this->result_);

                this->file_.close();
            }
        }
        for (int i = 0; i < this->slFiles_.size() + 1; i++)
            selection_->dynamicCall("moveDown()");
    }else qDebug() << "filesList isEmpty";
}

void PassportMakerTT::InsertImage()
{
    this->imagePath_ = this->directory_.absolutePath();
    this->imagePath_.append("/Безымянный.png");
    this->selection_->querySubObject("InlineShapes()")->dynamicCall("AddPicture(QString)", this->imagePath_);
//    QFile image(this->imagePath_);
//    image.remove();
}

void PassportMakerTT::SaveDocx()
{
    QString absolutePath;
    if (this->directory_.exists())
        absolutePath = this->directory_.absolutePath();

    QDir(absolutePath).mkdir("Паспорт");
    absolutePath.append("/Паспорт");
    absolutePath.append("/").append(this->executableFile_);
    absolutePath.remove(".txt");

    bool condition = false;
    while (!condition) {
        if (QFile(QString(absolutePath).append(".docx")).exists())
            absolutePath.append("_Newer");
        else condition = true;
    }
    absolutePath.append(".docx");
    this->activeDocument_->dynamicCall("SaveAs(FileName)", absolutePath);
    this->activeDocument_->dynamicCall("Close()");
}

void PassportMakerTT::SetTY(const QString &_TY)
{
    this->TY_ = _TY;
}

void PassportMakerTT::SetCorrectionTY(const QString &_correctionNumber)
{
    this->correctionTY_ = _correctionNumber;
}

void PassportMakerTT::SetKYnumber(const QString &_KYnumber)
{
    this->KYnumber_ = _KYnumber;
}

void PassportMakerTT::SetKYComment(const QString &_comment)
{
    this->commentKYConnection_ = _comment;
}

const QString &PassportMakerTT::getKYComment() const
{
    return this->commentKYConnection_;
}

void PassportMakerTT::SetDirPath(const QString &_dirPath)
{
    this->directory_.setPath(_dirPath);
    this->directory_.setFilter(QDir::Files);

    for (QString& format : this->slFileFormats_)
        if (!format.contains("*.")) format.push_front("*.");

    this->directory_.setNameFilters(this->slFileFormats_);
    if (!this->slFiles_.isEmpty())
        this->slFiles_.clear();
    for (QString name : this->directory_.entryList())
    {
        this->slFiles_.push_back(name);
        if (name.contains("txt"))
            this->executableFile_ = name;
    }
}

const QStringList &PassportMakerTT::GetSlExtensions() const
{
    return this->slFileFormats_;
}
