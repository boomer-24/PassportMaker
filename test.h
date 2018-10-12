#ifndef TEST_H
#define TEST_H

#include <QString>
#include <QStringList>

class Test
{
public:
    Test();
    Test(const QString& _name, const QString& _designation,
         const QString& _minValue, const QString& _maxValue,
         const QString& _condition,  const QString& _deviation);
    Test(const QString& _name, const QString& _designation,
         const QString& _minValue, const QString& _maxValue,
         const QString& _condition);

    QStringList GetTest() const;

    QString name() const;
    void setName(const QString &name);

    QString designation() const;
    void setDesignation(const QString &designation);

    QString minValue() const;
    void setMinValue(const QString &minValue);

    QString maxValue() const;
    void setMaxValue(const QString &maxValue);

    QString conditions() const;
    void setConditions(const QString &conditions);

    QString deviation() const;
    void setDeviation(const QString &deviation);

private:
    QStringList test_;
    QString name_;
    QString designation_;
    QString minValue_;
    QString maxValue_;
    QString conditions_;
    QString deviation_;
};

#endif // TEST_H
