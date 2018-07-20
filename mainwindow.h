#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QSortFilterProxyModel>
#include <QFileDialog>
#include <QCheckBox>
#include <QLabel>
#include <QLineEdit>
#include <QGridLayout>
#include <QTableWidget>
#include <QFileDialog>
#include <QAxObject>
#include <QPushButton>
#include <qDebug>
#include <QMessageBox>
#include <QCheckBox>
#include <QComboBox>
#include <QList>
#include <QTextCodec>
class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = 0);
    ~MainWindow();

private slots:
    void on_m_chooseBtn_clicked();
    void on_m_chooseoutBtn_clicked();
    void on_m_createBtn_clicked();
private:
    QPushButton *m_chooseBtn;
    QPushButton *m_chooseoutBtn;
    QPushButton *m_createBtn;
    QLineEdit *m_srcPath;
    QLineEdit *m_outPath;
    QLineEdit *m_colNum;

    QGridLayout *m_gbox;
    QList<QComboBox *> m_comboxlist;
};


#endif // MAINWINDOW_H
