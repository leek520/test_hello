#include "mainwindow.h"

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
{

    resize(500,200);
    QWidget *main = new QWidget(this);
    setCentralWidget(main);

    QVBoxLayout *vbox = new QVBoxLayout();
    main->setLayout(vbox);

    QHBoxLayout *hbox = new QHBoxLayout();
    QString currentPath = QDir::currentPath();
    //QString currentPath = "D:\\2-Work\\leek.project\\mam_tools\\xlsx\\菜单_6070BZ_E22.xls";
    m_srcPath = new QLineEdit(currentPath);
    m_chooseBtn = new QPushButton("选择");
    hbox->addWidget(new QLabel("表格路径："));
    hbox->addWidget(m_srcPath);
    hbox->addWidget(m_chooseBtn);
    hbox->setStretch(0,1);
    hbox->setStretch(1,5);
    hbox->setStretch(2,1);

    QHBoxLayout *hbox1 = new QHBoxLayout();
    QString outPath = QDir::currentPath();

    m_outPath = new QLineEdit(outPath);
    m_chooseoutBtn = new QPushButton("选择");
    hbox1->addWidget(new QLabel("输出路径："));
    hbox1->addWidget(m_outPath);
    hbox1->addWidget(m_chooseoutBtn);
    hbox1->setStretch(0,1);
    hbox1->setStretch(1,5);
    hbox1->setStretch(2,1);

    m_colNum = new QLineEdit("1");
    m_createBtn = new QPushButton("确定");
    QHBoxLayout *hbox_btn = new QHBoxLayout();
    hbox_btn->addWidget(new QLabel("语言数量："));
    hbox_btn->addWidget(m_colNum);
    hbox_btn->addWidget(m_createBtn);
    hbox_btn->setStretch(0,1);
    hbox_btn->setStretch(1,5);
    hbox_btn->setStretch(2,1);


    vbox->addLayout(hbox);
    vbox->addLayout(hbox1);
    vbox->addLayout(hbox_btn);

    connect(m_chooseBtn, SIGNAL(clicked(bool)), this, SLOT(on_m_chooseBtn_clicked()));
    connect(m_chooseoutBtn, SIGNAL(clicked(bool)), this, SLOT(on_m_chooseoutBtn_clicked()));
    connect(m_createBtn, SIGNAL(clicked(bool)), this, SLOT(on_m_createBtn_clicked()));

}

MainWindow::~MainWindow()
{

}
void MainWindow::on_m_chooseBtn_clicked()
{
    QString strFile = QFileDialog::getOpenFileName(this,QStringLiteral("选择Excel文件"),"",tr("Exel file(*.xls *.xlsx)"));
    m_srcPath->setText(strFile);
}

void MainWindow::on_m_chooseoutBtn_clicked()
{
    QString strFile = QFileDialog::getExistingDirectory(this,QStringLiteral("选择输出路径"),"");
    m_outPath->setText(strFile);
}

void MainWindow::on_m_createBtn_clicked()
{
    QStringList dirlist;
    dirlist<<"简体中文"<<"英文"<<"西班牙文"<<"繁体中文"<<"法文"<<"德文";
    QStringList suflist;
    suflist<<""<<" - ENG"<<" - 西班牙"<<" - 繁体"<<" - 法文"<<" - 德文";
    QStringList namelist;
    namelist<<"用户参数"<<"时间参数"<<"厂家参数"<<"运行参数"<<"预置变频器"
            <<"校准参数"<<"主界面参数及其它"<<"按钮输入输出端子功能"<<"压力温度切换"
            <<"运行状态"<<"故障"<<"预警"<<"硬件";

    int curCol = 0;

    QString strFile = m_srcPath->text();
    if (!strFile.endsWith(".xls") && !strFile.endsWith(".xlsx"))
    {
        return;
    }

    QString outDir = m_outPath->text();
    QDir pathDir(outDir);
    if (!pathDir.exists())
    {
        return;
    }

    //1、打开excel
    QAxObject excel("Excel.Application"); //加载Excel驱动
    excel.setProperty("Visible", false); //不显示Excel界面，如果为true会看到启动的Excel界面
    QAxObject* pWorkBooks = excel.querySubObject("WorkBooks");
    pWorkBooks->dynamicCall("Open (const QString&)", strFile);//打开指定文
    QAxObject* pWorkBook = excel.querySubObject("ActiveWorkBook");
    QAxObject* pWorkSheets = pWorkBook->querySubObject("Sheets");//获取工作表
    int nSheetCount = pWorkSheets->property("Count").toInt();  //获取工作表的数目

    for (int i=1; i<=nSheetCount; i++){
        QAxObject* pWorkSheet = pWorkBook->querySubObject("Sheets(int)", i);//获取第一张表
        QString name = pWorkSheet->property("Name").toString();
        name = name.trimmed();
        QAxObject *pUsedrange = pWorkSheet->querySubObject("UsedRange");//获取该sheet的使用范围对象
        QAxObject *pRows = pUsedrange->querySubObject("Rows");
        QAxObject *pColumns = pUsedrange->querySubObject("Columns");
        /*获取行数和列数*/
        int intCols = pColumns->property("Count").toInt();
        int intRows = pRows->property("Count").toInt();
        int intColStart = pColumns->property("Column").toInt();
        int intRowStart = pRows->property("Row").toInt();

        curCol = m_colNum->text().toInt();
        if (curCol < 0 || curCol > intCols )
        {
            qDebug()<<name<<intColStart<<intCols<<curCol<<"Error col num";
            continue;
        }
        for (int t=0; t<curCol; t++)
        {
            //2、创建txt
            QString path = QString("%1\/%2").arg(outDir).arg(dirlist[t]);
            QDir dir;
            if (!dir.exists(path))
            {
                dir.mkpath(path);
            }
            //QString filename = QString("%1\/%2\/%3.txt").arg(dir.currentPath()).arg(path).arg(name);
            QString fullname = namelist[i-1] + suflist[t];
            if (fullname.contains("用户参数") && t==1)
            {
                fullname = "用户参数 -ENG";
            }
            if (fullname.contains("硬件") && t==2)
            {
                fullname = "硬件 -繁体";
            }
            QString filename = QString("%1\/%2.txt").arg(path).arg(fullname);
            QFile file(filename);
            if(!file.open(QIODevice::WriteOnly|QIODevice::Text)){
                return;
            }
            QTextStream out(&file);
            out.setCodec(QTextCodec::codecForName("unicode"));//unicode小端模式
            out.setAutoDetectUnicode(true); //好像没用处
            QChar head = 0xfeff;//unicode文件头 文本里前两个字节为FFFE
            out << head;
            for (int j=intRowStart; j<=intRows; j++){
                QAxObject *range = pWorkSheet->querySubObject("Cells(int,int)", j, t+2); //获取cell的值
                QString value = range->dynamicCall("Value2()").toString();

                qDebug()<<value.toStdString().c_str();

                //3、插入txt
                out << value << "\n";
            }
            //4、关闭txt文件
            file.close();
        }
    }
    pWorkBooks->dynamicCall("Close()");

    QMessageBox::information(this, "信息", "已完成！");

}


