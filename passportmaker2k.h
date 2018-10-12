#ifndef PASPORTMAKER2K_H
#define PASPORTMAKER2K_H

#include <QString>
#include <QFile>
#include <QtXml>
#include <QAxObject>
#include <QCryptographicHash>

class PassportMaker2K
{
private:
    QString approveTrif_;
    QString approveShat_;
    QString approveGroupChief_1_;
    QString approveGroupChief_2_;
    QString title_;
    QString titlePasport_;
    QString titleOther_;
    QString firstParagraph1stString_;
    QStringList slElements_;
    QString firstParagraph2ndString_;
    QString TY_;
    QString correctionTY_;
    QString otherOnThe1stList_;
    QString description_, description1_;
    QString instruction_, instruction1_;
    QString guide_, guide1_;
    QString connectionInfoPinH_, connectionInfoPinS_;
    QString adapterName_, adapterNumber_, KYname_, KYnumber_;
    QVector<QPair<QString, QString>> vJPos_;
    int tableNumber_ = 0;

    QString selectedMode_;
    bool PIN_H_;

    QString executableFile_;

    QString selectedChief_;
    QString selectedAnimal_;
    QString selectedPosition_;

    QAxObject* wordApplication_;
    QAxObject* wordDocument_;
    QAxObject* activeDocument_;
    QAxObject* selection_;
    QAxObject* range_;
    QAxObject* tables_;
    QAxObject* table_;
    QAxObject* font_;
    QAxObject* paragraphFormat_;
    QAxObject* indent_;
    QAxObject* parForm_;

    QCryptographicHash* cryptHash_;
    QDir directory_;
    QStringList slFiles_;
    QStringList slFileFormats_;
    QFile file_;
    QByteArray data_;
    QString result_;

public:
    PassportMaker2K();
    ~PassportMaker2K();

    void Initialize(const QString& _pathToXML);
    void InitializeElementsSl(const QStringList& _slElements);
    void InitializeFileFormatSl(const QStringList& _slFileFormats);
    void CreateDocument();
    const QString GetWordVersion() const;

    bool ContainsElement(const QString& _element);
    void AddToElementsStringList(const QString& _element);
    void RemoveSpecificElement(const QString& _element);

    bool ContainsFormat(const QString& _format);
    void AddToFormatStringList(const QString& _format);
    void RemoveSpecificFormat(const QString& _format);
    void ClearFormatsStringList();

    int GetSlElementsSize() const;
    int GetSlFormatsSize() const;

    void TitleListCreate();
    void SecondListCreate();
    void OneMoreListCreate();
    void TestersList();
    void AddHashSheet();    
    void SaveDocx();

    void SetGroupChief(const QString &_chiefName);
    void SetAnimal(const QString& _animal);
    void SetAnimalPosition(const QString& _position);

    void SetTY(const QString& _TY);
    void SetCorrectionTY(const QString& _correctionNumber);
    void SetAdapterName(const QString& _adapterName);
    const QString& GetAdapterName();
    void SetAdapterNumber(const QString& _adapterNumber);
    const QString& GetAdapterNumber();
    void SetKYname(const QString& _KYname);
    const QString& GetKYname();
    void SetKYnumber(const QString& _KYnumber);
    const QString& GetKYnumber();
    void SetDirPath(const QString& _dirPath);
    void SetMode(const QString& _mode);
    void SetPinH(const bool& _pin);
    const QStringList& GetSlFormat() const;
    void SetVJPos(const QVector<QPair<QString, QString>> &_vJPos);
};

#endif // PASPORTMAKER2K_H
