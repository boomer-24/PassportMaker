#include "PassportMaker2K.h"

PassportMaker2K::PassportMaker2K()
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

PassportMaker2K::~PassportMaker2K()
{       
    this->wordApplication_->dynamicCall("Quit()");
    delete this->wordApplication_;
    qDebug() << "dTor";
}

void PassportMaker2K::Initialize(const QString& _pathToXML)
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
                        else if(domElement.tagName() == "title_2K") this->title_ = domElement.text();
                        else if(domElement.tagName() == "titlePasport_2K") this->titlePasport_ = domElement.text();
                        else if(domElement.tagName() == "titleOther_2K") this->titleOther_ = domElement.text();
                        else if(domElement.tagName() == "firstParagraph1stString_2K") this->firstParagraph1stString_ = domElement.text();
                        else if(domElement.tagName() == "firstParagraph2ndString_2K") this->firstParagraph2ndString_ = domElement.text();
                        else if(domElement.tagName() == "otherOnTheFirstList_2K") this->otherOnThe1stList_ = domElement.text();
                        else if(domElement.tagName() == "description_2K") this->description_ = domElement.text();
                        else if(domElement.tagName() == "description1_2K") this->description1_ = domElement.text();
                        else if(domElement.tagName() == "instruction_2K") this->instruction_ = domElement.text();
                        else if(domElement.tagName() == "instruction1_2K") this->instruction1_ = domElement.text();
                        else if(domElement.tagName() == "guide_2K") this->guide_ = domElement.text();
                        else if(domElement.tagName() == "guide1_2K") this->guide1_ = domElement.text();
                        else if(domElement.tagName() == "connectionInfoPinH_2K") this->connectionInfoPinH_ = domElement.text();
                        else if(domElement.tagName() == "connectionInfoPinS_2K") this->connectionInfoPinS_ = domElement.text();
                        //                        else if(domElement.tagName() == "adapterName_2K") this->adapterName_ = domElement.text();
                        //                        else if(domElement.tagName() == "adapterNumber_2K") this->adapterNumber_ = domElement.text();
                        else qDebug() << "Tag ne find";
                    }
                    domNode = domNode.nextSibling();
                }
            }
        } else qDebug() << "It`s no XML!";
    } else qDebug() << "File not open =/";
}

void PassportMaker2K::InitializeElementsSl(const QStringList &_slElements)
{
    this->slElements_ = _slElements;
}

void PassportMaker2K::InitializeFileFormatSl(const QStringList &_slFileFormats)
{
    this->slFileFormats_ = _slFileFormats;
}

void PassportMaker2K::CreateDocument()
{    
    this->TitleListCreate();
    this->SecondListCreate();
    this->OneMoreListCreate();
    this->TestersList();
    this->AddHashSheet();
}

const QString PassportMaker2K::GetWordVersion() const
{
    QVariant v = this->wordApplication_->dynamicCall("Version()");
    qDebug() << v.toString();
    return v.toString();
}

void PassportMaker2K::AddToElementsStringList(const QString &_element) // НЕ НАДО ЖЕTT
{
    this->slElements_.push_back(_element);
}

void PassportMaker2K::RemoveSpecificElement(const QString& _element)
{
    this->slElements_.removeOne(_element);
}

void PassportMaker2K::AddToFormatStringList(const QString& _format)
{
    this->slFileFormats_.push_back(_format.toLower());
}

void PassportMaker2K::ClearFormatsStringList()
{
    this->slFileFormats_.clear();
}

void PassportMaker2K::RemoveSpecificFormat(const QString& _format)
{
    //    QString format = QString("*.").append(_format);
    this->slFileFormats_.removeOne(_format);
}

bool PassportMaker2K::ContainsFormat(const QString& _format)
{
    if (this->slFileFormats_.contains(_format)) return true;
    else return false;
}

bool PassportMaker2K::ContainsElement(const QString& _element)
{
    if (this->slElements_.contains(_element)) return true;
    else return false;
}

int PassportMaker2K::GetSlElementsSize() const
{
    int size = this->slElements_.size();
    //    qDebug() << /*"__" << size << "__el" << */this->slElements2K_.size();
    return size;
}

int PassportMaker2K::GetSlFormatsSize() const
{
    int size = this->slFileFormats_.size();
    return size;
}

void PassportMaker2K::TitleListCreate()
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

void PassportMaker2K::SecondListCreate()
{    
    this->paragraphFormat_->dynamicCall("Alignment", "wdAlignParagraphJustify");

    this->selection_->querySubObject("PageSetup()")->setProperty("LeftMargin", 75);    // так работает
    this->paragraphFormat_->setProperty("LeftIndent", 0);                             // и так =)
    this->paragraphFormat_->setProperty("FirstLineIndent", 0);                             // и так =)

    this->selection_->dynamicCall("TypeText(const QString&)", this->firstParagraph1stString_);

    QString ims;
    for (int C = 0; C < this->slElements_.size(); C++)
    {
        ims.append(this->slElements_.at(C));
        if (C < this->slElements_.size() - 1)
            ims.append(", ");
    }

    QAxObject* section = this->activeDocument_->querySubObject("Sections(Index)", 1);
    QAxObject* footers = section->querySubObject("Footers(WdHeaderFooterIndex)", "wdHeaderFooterPrimary");
    QAxObject* rangeFoot = footers->querySubObject("Range()");
    rangeFoot->dynamicCall("InsertAfter(Text)", QString(ims).append("                               Разработчик:    ").append(this->selectedAnimal_));

    QAxObject* pageNumbers = this->activeDocument_->querySubObject("Sections(Index)", 1)->querySubObject("Footers(WdHeaderFooterIndex)", "wdHeaderFooterPrimary")->querySubObject("PageNumbers()")->querySubObject("Add(PageNumberAlignment, FirstPage)", "wdAlignPageNumberCenter", false);

    font_->setProperty("Bold", true);
    selection_->dynamicCall("TypeText(const QString&)",ims);
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

    parForm_->setProperty("LineSpacing", 12);

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

void PassportMaker2K::OneMoreListCreate()
{
    font_->setProperty("Bold", true);
    QString tableName("Таблица ");
    tableName.append(QString::number(++this->tableNumber_)).append(". Перечень проверяемых параметров и отклонений от ТУ.");
    selection_->dynamicCall("TypeText(const QString&)", tableName);
    font_->setProperty("Bold", false);
    range_ = selection_->querySubObject("Range()");
    table_ = tables_->querySubObject("Add(Range,NumRows,NumColumns, DefaultTableBehavior, AutoFitBehavior)",range_->asVariant(), 2, 5, 1, 2);

    selection_->dynamicCall("moveDown()");
    selection_->dynamicCall("moveDown()");
    selection_->dynamicCall("moveDown()");
    selection_->dynamicCall("TypeParagraph()");

    QAxObject* cell_11 = table_->querySubObject("Cell(Row,  Column)", 1, 1);
    cell_11->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "4.4", "wdAdjustNone");
    QAxObject* rangeCell = cell_11->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Наименование параметра, единица измерения. Нумерация по ТУ");
    QAxObject* cell_21 = table_->querySubObject("Cell(Row,  Column)", 2, 1);
    cell_21->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "4.4", "wdAdjustNone");

    QAxObject* cell_12 = table_->querySubObject("Cell(Row,  Column)", 1, 2);
    cell_12->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "2.8", "wdAdjustNone");
    rangeCell = cell_12->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Обозначение");
    QAxObject* cell_22 = table_->querySubObject("Cell(Row,  Column)", 2, 2);
    cell_22->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "2.8", "wdAdjustNone");

    QAxObject* cell_13 = table_->querySubObject("Cell(Row,  Column)", 1, 3);
    cell_13->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "2.8", "wdAdjustNone");
    rangeCell = cell_13->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Норма");
    QAxObject* cell_23 = table_->querySubObject("Cell(Row,  Column)", 2, 3);
    cell_23->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "2.8", "wdAdjustNone");

    QAxObject* cell_14 = table_->querySubObject("Cell(Row,  Column)", 1, 4);
    cell_14->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "3.73", "wdAdjustNone");
    rangeCell = cell_14->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Условия");
    QAxObject* cell_24 = table_->querySubObject("Cell(Row,  Column)", 2, 4);
    cell_24->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "3.73", "wdAdjustNone");

    QAxObject* cell_15 = table_->querySubObject("Cell(Row,  Column)", 1, 5);
    cell_15->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "3.5", "wdAdjustNone");
    rangeCell = cell_15->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Отклонения от ТУ");
    QAxObject* cell_25 = table_->querySubObject("Cell(Row,  Column)", 2, 5);
    cell_25->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "3.5", "wdAdjustNone");

    QAxObject* cella = table_->querySubObject("Cell(Row,  Column)", 1, 3);
    cella->dynamicCall("Split(NumRows, NumColumns)", 2, 1);
    QAxObject* cella1 = table_->querySubObject("Cell(Row,  Column)", 2, 3);
    cella1->dynamicCall("Split(NumRows, NumColumns)", 1, 2);
    QAxObject* cella2 = table_->querySubObject("Cell(Row,  Column)", 3, 3);
    cella2->dynamicCall("Split(NumRows, NumColumns)", 1, 2);

    QAxObject* cell_23_ = table_->querySubObject("Cell(Row,  Column)", 2, 3);
    rangeCell = cell_23_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Не менее");
    QAxObject* cell_24_ = table_->querySubObject("Cell(Row,  Column)", 2, 4);
    rangeCell = cell_24_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Не более");

    font_->setProperty("Bold", true);
    tableName = "Таблица ";
    tableName.append(QString::number(++this->tableNumber_)).append(". Перечень непроверяемых параметров.");
    selection_->dynamicCall("TypeText(const QString&)", tableName);
    font_->setProperty("Bold", false);
    range_ = selection_->querySubObject("Range()");
    table_ = tables_->querySubObject("Add(Range,NumRows,NumColumns, DefaultTableBehavior, AutoFitBehavior)",range_->asVariant(), 2, 4, 1, 2);

    selection_->dynamicCall("moveDown()");
    selection_->dynamicCall("moveDown()");
    selection_->dynamicCall("moveDown()");
    selection_->dynamicCall("TypeParagraph()");

    QAxObject* ccell11 = table_->querySubObject("Cell(Row,  Column)", 1, 1);
    ccell11->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "4.4", "wdAdjustNone");
    ccell11->querySubObject("Range()")->dynamicCall("InsertAfter(Text)", "Наименование параметра, единица измерения.");
    QAxObject* ccell21 = table_->querySubObject("Cell(Row,  Column)", 2, 1);
    ccell21->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "4.4", "wdAdjustNone");

    QAxObject* ccell12 = table_->querySubObject("Cell(Row,  Column)", 1, 2);
    ccell12->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "2.8", "wdAdjustNone");
    ccell12->querySubObject("Range()")->dynamicCall("InsertAfter(Text)", "Обозначение");
    QAxObject* ccell22 = table_->querySubObject("Cell(Row,  Column)", 2, 2);
    ccell22->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "2.8", "wdAdjustNone");

    QAxObject* ccell13 = table_->querySubObject("Cell(Row,  Column)", 1, 3);
    ccell13->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "2.8", "wdAdjustNone");
    ccell13->querySubObject("Range()")->dynamicCall("InsertAfter(Text)", "Норма");
    QAxObject* ccell23 = table_->querySubObject("Cell(Row,  Column)", 2, 3);
    ccell23->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "2.8", "wdAdjustNone");

    QAxObject* ccell14 = table_->querySubObject("Cell(Row,  Column)", 1, 4);
    ccell14->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "7.23", "wdAdjustNone");
    ccell14->querySubObject("Range()")->dynamicCall("InsertAfter(Text)", "Условия");
    QAxObject* ccell24 = table_->querySubObject("Cell(Row,  Column)", 2, 4);
    ccell24->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "7.23", "wdAdjustNone");

    QAxObject* ccella = table_->querySubObject("Cell(Row,  Column)", 1, 3);
    ccella->dynamicCall("Split(NumRows, NumColumns)", 2, 1);
    QAxObject* ccella1 = table_->querySubObject("Cell(Row,  Column)", 2, 3);
    ccella1->dynamicCall("Split(NumRows, NumColumns)", 1, 2);
    QAxObject* ccella2 = table_->querySubObject("Cell(Row,  Column)", 3, 3);
    ccella2->dynamicCall("Split(NumRows, NumColumns)", 1, 2);

    QAxObject* ccell_23_ = table_->querySubObject("Cell(Row,  Column)", 2, 3);
    rangeCell = ccell_23_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Не менее");
    QAxObject* ccell_24_ = table_->querySubObject("Cell(Row,  Column)", 2, 4);
    rangeCell = ccell_24_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Не более");

    selection_->dynamicCall("moveDown()");
    selection_->dynamicCall("moveDown()");
    //    selection2K_->dynamicCall("moveDown()");
    //    selection2K_->dynamicCall("TypeParagraph()");

    font_->setProperty("Bold", true);
    tableName = "Таблица ";
    tableName.append(QString::number(++this->tableNumber_)).append(". Оснастка.");
    selection_->dynamicCall("TypeText(const QString&)", tableName);
    font_->setProperty("Bold", false);
    range_ = selection_->querySubObject("Range()");
    table_ = tables_->querySubObject("Add(Range,NumRows,NumColumns, DefaultTableBehavior, AutoFitBehavior)",range_->asVariant(), 3, 3, 1, 2);

    selection_->dynamicCall("TypeText(const QString&)","Тип оснастки");
    selection_->dynamicCall("moveRight()");
    selection_->dynamicCall("TypeText(const QString&)","Название оснастки");
    selection_->dynamicCall("moveRight()");
    selection_->dynamicCall("TypeText(const QString&)","Регистрационный номер оснастки");
    selection_->dynamicCall("moveDown()");
    selection_->dynamicCall("moveLeft()");
    selection_->dynamicCall("moveLeft()");
    selection_->dynamicCall("TypeText(const QString&)","Адаптер");
    selection_->dynamicCall("moveDown()");
    selection_->dynamicCall("TypeText(const QString&)","КУ");
    selection_->dynamicCall("moveRight()");
    selection_->dynamicCall("moveUp()");
    selection_->dynamicCall("TypeText(const QString&)",this->adapterName_);
    selection_->dynamicCall("moveRight()");
    selection_->dynamicCall("TypeText(const QString&)",this->adapterNumber_);
    selection_->dynamicCall("moveDown()");
    selection_->dynamicCall("moveLeft()");
    selection_->dynamicCall("TypeText(const QString&)",this->KYname_);
    selection_->dynamicCall("moveRight()");
    selection_->dynamicCall("TypeText(const QString&)",this->KYnumber_);

    selection_->dynamicCall("moveDown()");
    selection_->dynamicCall("moveDown()");
    selection_->dynamicCall("TypeParagraph()");

    if (!this->vJPos_.isEmpty())
    {
        font_->setProperty("Bold", true);
        tableName = "Таблица ";
        tableName.append(QString::number(++this->tableNumber_)).append(". Конфигурация перемычек.");
        selection_->dynamicCall("TypeText(const QString&)", tableName);
        font_->setProperty("Bold", false);
        range_ = selection_->querySubObject("Range()");
        table_ = tables_->querySubObject("Add(Range,NumRows,NumColumns, DefaultTableBehavior, AutoFitBehavior)",
                                         range_->asVariant(), 2, this->vJPos_.size() + 1, 1, 2);
        QAxObject *cell, *rangeCell;
        cell = table_->querySubObject("Cell(Row, Column)", 1, 1);
        rangeCell = cell->querySubObject("Range()");
        rangeCell->dynamicCall("InsertAfter(Text)", "№ перемычки");
        cell = table_->querySubObject("Cell(Row, Column)", 2, 1);
        rangeCell = cell->querySubObject("Range()");
        rangeCell->dynamicCall("InsertAfter(Text)", "Положение");
        for (int i = 0; i < this->vJPos_.size(); i++)
        {
            cell = table_->querySubObject("Cell(Row, Column)", 1, i + 2);
            rangeCell = cell->querySubObject("Range()");
            rangeCell->dynamicCall("InsertAfter(Text)", this->vJPos_.at(i).first);

            cell = table_->querySubObject("Cell(Row, Column)", 2, i + 2);
            rangeCell = cell->querySubObject("Range()");
            rangeCell->dynamicCall("InsertAfter(Text)", this->vJPos_.at(i).second);
        }

        for (int i = 0; i < 10; i++)
            selection_->dynamicCall("moveDown()");
        //    selection2K_->dynamicCall("TypeText(const QString&)","При использовании адаптера Adap60 перемычки не устанавливаются.");
        selection_->dynamicCall("TypeParagraph()");
        selection_->dynamicCall("TypeParagraph()");
    }


    this->FillTableConnection();

    selection_->dynamicCall("InsertBreak()", "wdPageBreak"); //новый лист
}

void PassportMaker2K::FillTableConnection()
{
    font_->setProperty("Bold", true);
    QString tableName = "Таблица ";
    tableName.append(QString::number(++this->tableNumber_)).append(". Подключение выводов адаптера к тестеру.");
    selection_->dynamicCall("TypeText(const QString&)", tableName);
    font_->setProperty("Bold", false);
    range_ = selection_->querySubObject("Range()");
    if (!this->vvTableConnection_.isEmpty())
    {
        table_ = tables_->querySubObject("Add(Range,NumRows,NumColumns, DefaultTableBehavior, AutoFitBehavior)",
                                         range_->asVariant(), this->vvTableConnection_.size(), this->vvTableConnection_.first()/*.at(0)*/.size(), 1, 2);

        //        QAxObject* rows = table_->querySubObject("Rows()");
        //        qDebug() << "rows" << rows->asVariant().toString();   //НЕ РАБОТАЕТ ТАК
        //        rows->setProperty("Alignment", "wdAlignRowCenter");
        for (int row = 0; row < this->vvTableConnection_.size(); row++)
        {
            for (int column = 0; column < this->vvTableConnection_.at(row).size(); column++)
            {
                if (row == 0 && column == 0)
                {
                    QAxObject* cell11 = table_->querySubObject("Cell(Row, Column)", row + 1, column + 1);
                    QAxObject* borders = cell11->querySubObject("Borders()");
                    QAxObject* border = borders->querySubObject("Item(WdBorderType)", "wdBorderDiagonalDown");
                    border->setProperty("LineStyle", "wdLineStyleSingle");
                    QAxObject* rangeCell = cell11->querySubObject("Range()");
                    rangeCell->dynamicCall("InsertAfter(Text)", "    Тестер\n\n\nАдаптер");
                    selection_->dynamicCall("moveDown()");
                    selection_->dynamicCall("moveDown()");
                    selection_->dynamicCall("moveDown()");
                    selection_->dynamicCall("moveDown()");
                }
                QAxObject* cell = table_->querySubObject("Cell(Row,  Column)", row + 1, column + 1);
                QAxObject* rangeCell = cell->querySubObject("Range()");
                rangeCell->dynamicCall("InsertAfter(Text)", this->vvTableConnection_.at(row).at(column));
            }
            selection_->dynamicCall("moveDown()");
        }
        //        for (int i = 0; i < this->vvStrAndTableConnection_.size() + 10; i++)
    }
}

void PassportMaker2K::TestersList()
{
    range_ = selection_->querySubObject("Range()");
    QAxObject* ParForm = this->range_->querySubObject("ParagraphFormat()");
    ParForm->setProperty("SpaceAfter", 0);

    font_->setProperty("Bold", true);
    selection_->dynamicCall("TypeText(const QString&)","Лист испытателя");   selection_->dynamicCall("TypeParagraph()");

    font_->setProperty("Bold", false);
    selection_->dynamicCall("TypeText(const QString&)", "Изделие: ");
    font_->setProperty("Bold", true);
    QString imss;
    for (int i = 0; i < this->slElements_.size(); i++)
    {
        imss.append(this->slElements_.at(i));
        if (i < this->slElements_.size() - 1) imss.append(", ");
    }
    selection_->dynamicCall("TypeText(const QString&)", imss);  selection_->dynamicCall("TypeParagraph()");

    font_->setProperty("Bold", false);
    selection_->dynamicCall("TypeText(const QString&)", "ТУ: ");
    font_->setProperty("Bold", true);
    selection_->dynamicCall("TypeText(const QString&)", this->TY_.append(", коррекция ").append(this->correctionTY_));   selection_->dynamicCall("TypeParagraph()");

    font_->setProperty("Bold", false);
    selection_->dynamicCall("TypeText(const QString&)", "Оболочка: ");
    font_->setProperty("Bold", true);
    selection_->dynamicCall("TypeText(const QString&)", this->selectedMode_);
    if (!this->softwareVersion_.isEmpty())
    {
        font_->setProperty("Bold", false);
        selection_->dynamicCall("TypeText(const QString&)", " версия: ");
        font_->setProperty("Bold", true);
        selection_->dynamicCall("TypeText(const QString&)", this->softwareVersion_);
    }
    /*selection2K_->dynamicCall("TypeParagraph()");*/

    font_->setProperty("Bold", false);   selection_->dynamicCall("TypeParagraph()");
    selection_->dynamicCall("TypeText(const QString&)", "Адаптер универсальный: ");
    font_->setProperty("Bold", true);
    selection_->dynamicCall("TypeText(const QString&)", this->adapterNumber_);
    //    selection2K_->dynamicCall("TypeText(const QString&)", this->adapterName_.append("    ").append(this->adapterNumber_));
    selection_->dynamicCall("TypeParagraph()");

    font_->setProperty("Bold", false);
    selection_->dynamicCall("TypeText(const QString&)", "Контактирующее устройство: ");
    font_->setProperty("Bold", true);
    selection_->dynamicCall("TypeText(const QString&)", this->KYname_.append("    ").append(this->KYnumber_));
    selection_->dynamicCall("TypeParagraph()");
    font_->setProperty("Bold", false);

    if (!this->vJPos_.isEmpty())
    {
        QString jumperPos = "Установить перемычки";
        for (int i = 0; i < this->vJPos_.size(); i++)
        {
            const QPair<QString, QString> &rPair = this->vJPos_.at(i);
            if (rPair.second.contains("Не устан") || rPair.second.isEmpty())
                continue;
            jumperPos.append(" ").append(rPair.first);
            if (rPair.second.contains("Устан"))
                jumperPos.append(",");
            else
            {
                jumperPos.append(" в положение ").append(rPair.second).append(",");
            }
        }
        if (jumperPos.at(jumperPos.size() - 1) == ",")
            jumperPos.remove(jumperPos.size() - 1, 1);
        selection_->dynamicCall("TypeText(const QString&)", jumperPos);  selection_->dynamicCall("TypeParagraph()");
        selection_->dynamicCall("TypeParagraph()");
    }
    //    selection_->dynamicCall("TypeText(const QString&)", "На адаптере Adap60 установить перемычку J1.");

    font_->setProperty("Bold", true);
    selection_->dynamicCall("TypeText(const QString&)", "Подключение адаптеров для проверки изделия: "); selection_->dynamicCall("TypeParagraph()");
    font_->setProperty("Bold", false);
    selection_->dynamicCall("TypeText(const QString&)", "Запускать файл ");
    font_->setProperty("Bold", true);
    selection_->dynamicCall("TypeText(const QString&)", this->executableFile_);
    font_->setProperty("Bold", false);

    selection_->querySubObject("InlineShapes")->dynamicCall("AddHorizontalLineStandard()");
    font_->setProperty("Size", 12);    font_->setProperty("Name", "Times New Roman");



    selection_->dynamicCall("TypeText(const QString&)", this->connectionInfo2K_);
    //    if (this->PIN_H_)
    //        selection_->dynamicCall("TypeText(const QString&)", this->connectionInfoPinH_);
    //    else
    //        selection_->dynamicCall("TypeText(const QString&)", this->connectionInfoPinS_);



    selection_->querySubObject("InlineShapes")->dynamicCall("AddHorizontalLineStandard()");

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
    rangeCell->dynamicCall("InsertAfter(Text)", QString("_______").append(this->selectedAnimal_));

    QAxObject* cell_21 = table_->querySubObject("Cell(Row,  Column)", 2, 1);
    cell_21->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "10.5", "wdAdjustNone");
    QAxObject* cell_22 = table_->querySubObject("Cell(Row,  Column)", 2, 2);
    cell_22->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "5.5", "wdAdjustNone");
    rangeCell = cell_22->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "_______Дата");

    selection_->dynamicCall("moveDown()");
    selection_->dynamicCall("moveDown()");
    selection_->dynamicCall("moveDown()");

    selection_->dynamicCall("TypeParagraph()");
    selection_->dynamicCall("TypeParagraph()");
}

void PassportMaker2K::AddHashSheet()
{
    if (!this->slFiles_.isEmpty())
    {
        font_->setProperty("Size", 12);
        font_->setProperty("Name", "Times New Roman");
        font_->setProperty("Bold", true);
        QString tableName = "Таблица ";
        tableName.append(QString::number(++this->tableNumber_)).append(". Контрольные суммы файлов программы (SHA-1).");
        selection_->dynamicCall("TypeText(const QString&)", tableName);
        font_->setProperty("Bold", false);
        font_->setProperty("Size", 11);
        range_ = selection_->querySubObject("Range()");
        table_ = tables_->querySubObject("Add(Range, NumRows, NumColumns, DefaultTableBehavior, AutoFitBehavior)",
                                         range_->asVariant(), this->slFiles_.size() + 1, 3, 1, 0);

        QAxObject* cellTop0 = table_->querySubObject("Cell(Row, Column)", 1, 1);
        QAxObject* cellTopRange0 = cellTop0->querySubObject("Range()");
        cellTopRange0->dynamicCall("InsertAfter(Text)", "№ Версии");

        QAxObject* cellTop1 = table_->querySubObject("Cell(Row, Column)", 1, 2);
        QAxObject* cellTopRange1 = cellTop1->querySubObject("Range()");
        cellTopRange1->dynamicCall("InsertAfter(Text)", "Файл");

        QAxObject* cellTop2 = table_->querySubObject("Cell(Row, Column)", 1, 3);
        QAxObject* cellTopRange2 = cellTop2->querySubObject("Range()");
        cellTopRange2->dynamicCall("InsertAfter(Text)", "Контрольная сумма");

        QAxObject* allign = selection_->querySubObject("ParagraphFormat()");

        for (int i = 0; i < this->slFiles_.size(); i++)
        {
            QAxObject* cell0 = table_->querySubObject("Cell(Row, Column)", i + 2, 1);
            QAxObject* cellRange0 = cell0->querySubObject("Range()");
            allign->dynamicCall("Alignment", "wdAlignParagraphCenter");
            cellRange0->dynamicCall("InsertAfter(Text)", "         1");
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

void PassportMaker2K::SetGroupChief(const QString &_chiefName)
{
    this->selectedChief_ = _chiefName;
}

void PassportMaker2K::SetAnimal(const QString& _animal)
{
    this->selectedAnimal_ = _animal;
}

void PassportMaker2K::SetAnimalPosition(const QString& _position)
{
    this->selectedPosition_ = _position;
}

void PassportMaker2K::SetTY(const QString& _TY)
{
    this->TY_ = _TY;
}

void PassportMaker2K::SetCorrectionTY(const QString& _correctionNumber)
{
    this->correctionTY_ = _correctionNumber;
}

void PassportMaker2K::SetAdapterName(const QString& _adapterName)
{
    this->adapterName_ = _adapterName;
}

const QString& PassportMaker2K::GetAdapterName()
{
    return this->adapterName_;
}

void PassportMaker2K::SetAdapterNumber(const QString& _adapterNumber)
{
    this->adapterNumber_ = _adapterNumber;
}

const QString& PassportMaker2K::GetAdapterNumber()
{
    return this->adapterNumber_;
}

void PassportMaker2K::SetKYname(const QString& _KYname)
{
    this->KYname_ = _KYname;
}

const QString& PassportMaker2K::GetKYname()
{
    return this->KYname_;
}

void PassportMaker2K::SetKYnumber(const QString& _KYnumber)
{
    this->KYnumber_ = _KYnumber;
}

const QString& PassportMaker2K::GetKYnumber()
{
    return this->KYnumber_;
}

void PassportMaker2K::SetDirPath(const QString& _dirPath)
{
    this->directory_.setPath(_dirPath);
    this->directory_.setFilter(QDir::Files);

    for (QString& format : this->slFileFormats_ )
        if (!format.contains("*.")) format.push_front("*.");

    this->directory_.setNameFilters(this->slFileFormats_);
    if (!this->slFiles_.isEmpty())
        this->slFiles_.clear();
    for (QString name : this->directory_.entryList())
    {
        this->slFiles_.push_back(name);
        if (name.contains("cpf"))
            this->executableFile_ = name;
    }
}

void PassportMaker2K::SetMode(const QString& _mode)
{
    this->selectedMode_ = _mode;
}

void PassportMaker2K::SetPinH(const bool& _pin)
{
    this->PIN_H_ = _pin;
}

void PassportMaker2K::SetVJPos(const QVector<QPair<QString, QString>> &_vJPos)
{
    this->vJPos_ = _vJPos;
}

void PassportMaker2K::SaveDocx()
{
    QString absolutePath;
    if (this->directory_.exists())
        absolutePath = this->directory_.absolutePath();

    QDir(absolutePath).mkdir("Паспорт");
    absolutePath.append("/Паспорт");
    absolutePath.append("/").append(this->executableFile_);
    absolutePath.remove(".cpf");

    bool condition = false;
    while (!condition) {
        if (QFile(QString(absolutePath).append(".docx")).exists())
            absolutePath.append("_создан_").append(QDateTime::currentDateTime().toString("yyyy.MM.dd-hh-mm-ss"));
        else condition = true;
    }
    absolutePath.append(".docx");
    this->activeDocument_->dynamicCall("SaveAs(FileName)", absolutePath);
    this->activeDocument_->dynamicCall("Close()");
}

const QStringList& PassportMaker2K::GetSlFormat() const
{
    return this->slFileFormats_;
}

void PassportMaker2K::setStrAndVvTableConnection(const QVector<QVector<QString> > &_vvTableConnection)
{
    vvTableConnection_ = _vvTableConnection;
    if (!_vvTableConnection.isEmpty())
    {
        for (int i = 1; i < _vvTableConnection.at(0).size(); i++) // тут проход по нулевой строке с номерами тестеров
        {
            QString testerNumber = _vvTableConnection.at(0).at(i);
            this->connectionInfo2K_.append("Formula2K №").append(testerNumber).append(":").append("\n");
            for (int j = 1; j < _vvTableConnection.size(); j++) // тут по всем строкам таблицы, кроме первой
            {
                QString strResult("Подключить кабель «");
                QString valueFirstColumn(_vvTableConnection.at(j).at(0));
                strResult.append(valueFirstColumn);
                strResult.append("» адаптера к разъёму «");
                QString valueInCross(_vvTableConnection.at(j).at(i));
                strResult.append(valueInCross);
                strResult.append("» Тестера.\n");
                this->connectionInfo2K_.append(strResult);
            }
        }
    }
}

void PassportMaker2K::setSoftwareVersion(const QString &_softwareVersion)
{
    this->softwareVersion_ = _softwareVersion;
}

