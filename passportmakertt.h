#ifndef PassportMakerTT_H
#define PassportMakerTT_H

#include "test.h"

#include <QObject>
#include <QString>
#include <QFile>
#include <QtXml>
#include <QAxObject>
#include <QCryptographicHash>

class PassportMakerTT : public QObject
{
    Q_OBJECT

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
    QString commentKYConnection_;
    QString guide_, guide1_;
    QString KYnumber_;
    QString imagePath_;

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

    QVector<Test> vVerifiableTests_;
    QVector<Test> vUnVerifiableTests_;

public:
    PassportMakerTT(QObject* _parent);
    ~PassportMakerTT();

    void Initialize(const QString& _pathToXML);
    void InitializeElementsSl(const QStringList& _slElements);
    void InitializeFileFormatSl(const QStringList& _slFileFormats);
    void CreateDocument();

    bool ContainsElement(const QString& _element);
    void AddToElementsStringList(const QString& _element);
    void RemoveSpecificElement(const QString& _element);

    bool ContainsFormat(const QString& _format);
    void AddToFormatsStringList(const QString& _format);
    void RemoveSpecificFormat(const QString& _format);

    int GetSlElementsSize() const;
    int GetSlFormatsSize() const;

    void SetGroupChief(const QString& _chiefName);
    void SetAnimal(const QString& _animalName);
    void SetAnimalPosition(const QString& _animalPosition);
    void PushVerifiableTest(const Test& _test);
    void PushUnVerifiableTest(const Test& _test);
    void SetVVerifiableTests(const QVector<Test> &vVerifiableTests);
    void SetVUnVerifiableTests(const QVector<Test> &vUnVerifiableTests);

    void TitleListCreate();
    void SecondListCreate();
    void OneMoreListCreate();
    void TestersList();
    void AddHashSheet();
    void InsertImage();
    void SaveDocx();

    void SetTY(const QString& _TY);
    void SetCorrectionTY(const QString& _correctionNumber);
    void SetKYnumber(const QString& _KYnumber);
    void SetKYComment(const QString& _comment);
    const QString& getKYComment() const;
    void SetDirPath(const QString& _dirPath);
    const QStringList& GetSlExtensions() const;
    QString getImagePath() const;

public slots:
    void slotCreateAndSavePassport();
};

#endif // PassportMakerTT_H
