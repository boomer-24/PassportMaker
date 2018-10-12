#include "test.h"

Test::Test()
{

}

Test::Test(const QString &_name, const QString &_designation, const QString &_minValue, const QString &_maxValue,
           const QString &_condition, const QString &_deviation)
{
    this->name_ = _name;
    this->test_.push_back(_name);
    this->designation_ = _designation;
    this->test_.push_back(_designation);
    this->minValue_ = _minValue;
    this->test_.push_back(_minValue);
    this->maxValue_ = _maxValue;
    this->test_.push_back(_maxValue);
    this->conditions_ = _condition;
    this->test_.push_back(_condition);
    this->deviation_ =_deviation;
    this->test_.push_back(_deviation);
}

Test::Test(const QString &_name, const QString &_designation, const QString &_minValue,
           const QString &_maxValue, const QString &_condition)
{
    this->name_ = _name;
    this->test_.push_back(_name);
    this->designation_ = _designation;
    this->test_.push_back(_designation);
    this->minValue_ = _minValue;
    this->test_.push_back(_minValue);
    this->maxValue_ = _maxValue;
    this->test_.push_back(_maxValue);
    this->conditions_ = _condition;
    this->test_.push_back(_condition);
}

QStringList Test::GetTest() const
{
    return this->test_;
}

QString Test::name() const
{
    return name_;
}

void Test::setName(const QString &name)
{
    name_ = name;
}

QString Test::designation() const
{
    return designation_;
}

void Test::setDesignation(const QString &designation)
{
    designation_ = designation;
}

QString Test::minValue() const
{
    return minValue_;
}

void Test::setMinValue(const QString &minValue)
{
    minValue_ = minValue;
}

QString Test::maxValue() const
{
    return maxValue_;
}

void Test::setMaxValue(const QString &maxValue)
{
    maxValue_ = maxValue;
}

QString Test::conditions() const
{
    return conditions_;
}

void Test::setConditions(const QString &conditions)
{
    conditions_ = conditions;
}

QString Test::deviation() const
{
    return deviation_;
}

void Test::setDeviation(const QString &deviation)
{
    deviation_ = deviation;
}
