#-------------------------------------------------
#
# Project created by QtCreator 2018-04-24T11:57:35
#
#-------------------------------------------------

QT       += core gui

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets xml

TARGET = Paspartu
TEMPLATE = app

win32: RC_ICONS = $$PWD/paspart.ico

CONFIG += qaxcontainer

# The following define makes your compiler emit warnings if you use
# any feature of Qt which as been marked as deprecated (the exact warnings
# depend on your compiler). Please consult the documentation of the
# deprecated API in order to know how to port your code away from it.
DEFINES += QT_DEPRECATED_WARNINGS

# You can also make your code fail to compile if you use deprecated APIs.
# In order to do so, uncomment the following line.
# You can also select to disable deprecated APIs only up to a certain version of Qt.
#DEFINES += QT_DISABLE_DEPRECATED_BEFORE=0x060000    # disables all the APIs deprecated before Qt 6.0.0


SOURCES += main.cpp\
        mainwindow.cpp \
    passportmakertt.cpp \
    testsdialog.cpp \
    test.cpp \
    jumpersdialog.cpp \
    connectiondialog.cpp \
    lastcheckdialog.cpp \
    passportmaker2k.cpp \
    softwareversiondialog.cpp

HEADERS  += mainwindow.h \
    passportmakertt.h \
    testsdialog.h \
    test.h \
    jumpersdialog.h \
    connectiondialog.h \
    lastcheckdialog.h \
    passportmaker2k.h \
    softwareversiondialog.h

FORMS    += mainwindow.ui \
    testsdialog.ui \
    jumpersdialog.ui \
    connectiondialog.ui \
    lastcheckdialog.ui \
    softwareversiondialog.ui

RESOURCES += \
    res.qrc
