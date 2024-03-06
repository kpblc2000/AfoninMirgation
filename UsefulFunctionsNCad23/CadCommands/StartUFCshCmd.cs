
#if NCAD
using HostMgd.ApplicationServices;
using HostMgd.EditorInput;
using Teigha.Runtime;
#elif ACAD
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;
#endif

namespace UsefulFunctionsNCad23.CadCommands
{
    public static class StartUFCshCmd
    {
        [CommandMethod("StartUFCsh", CommandFlags.Modal)]
        public static void StartUFCshCommand()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
#if NCAD
            ed.WriteMessage("Полезные функции стартуют без панелей и кнопок\n");
#else
            //_______процедура вставки пользовательского меню с панелями и кнопками
            RibbonButton ribbonButton1 = new RibbonButton();

            ribbonButton1.Id = "_ribbonButton1";
            ribbonButton1.Text = "Сделай_попер";
            ribbonButton1.ShowText = true;
            ribbonButton1.Tag= ribbonButton1.Id;
            // привязываем к кнопке обработчик нажатия
            ribbonButton1.CommandHandler = new CommandHandler_ribbonButton1();

            RibbonButton ribbonButton2 = new RibbonButton();
            ribbonButton2.Id = "_ribbonButton2";
            ribbonButton2.Text = "Создать .pop";
            ribbonButton2.ShowText = true;
            ribbonButton2.Tag = ribbonButton2.Id;
            // привязываем к кнопке обработчик нажатия
            ribbonButton2.CommandHandler = new CommandHandler_ribbonButton2();

            RibbonButton ribbonButton3 = new RibbonButton();
            ribbonButton3.Id = "_ribbonButton3";
            ribbonButton3.Text = "Длина ломаной";
            ribbonButton3.ShowText = true;
            ribbonButton3.Tag = ribbonButton3.Id;
            // привязываем к кнопке обработчик нажатия
            ribbonButton3.CommandHandler = new CommandHandler_ribbonButton3();

            RibbonButton ribbonButton4 = new RibbonButton();
            ribbonButton4.Id = "_ribbonButton4";
            ribbonButton4.Text = "Чертить .рор";
            ribbonButton4.ShowText = true;
            ribbonButton4.Tag = ribbonButton4.Id;
            // привязываем к кнопке обработчик нажатия
            ribbonButton4.CommandHandler = new CommandHandler_ribbonButton4();

            RibbonButton ribbonButton5 = new RibbonButton();
            ribbonButton5.Id = "_ribbonButton5";
            ribbonButton5.Text = "Настройки";
            ribbonButton5.ShowText = true;
            ribbonButton5.Tag = ribbonButton5.Id;
            // привязываем к кнопке обработчик нажатия
            ribbonButton5.CommandHandler = new CommandHandler_ribbonButton5();

            RibbonButton ribbonButton6 = new RibbonButton();
            ribbonButton6.Id = "_ribbonButton6";
            ribbonButton6.Text = "Линии на 0";
            ribbonButton6.ShowText = true;
            ribbonButton6.Tag = ribbonButton6.Id;
            // привязываем к кнопке обработчик нажатия
            ribbonButton6.CommandHandler = new CommandHandler_ribbonButton6();

            RibbonButton ribbonButton7 = new RibbonButton();
            ribbonButton7.Id = "_ribbonButton7";
            ribbonButton7.Text = "Описание хар.линии";
            ribbonButton7.ShowText = true;
            ribbonButton7.Tag = ribbonButton7.Id;
            // привязываем к кнопке обработчик нажатия
            ribbonButton7.CommandHandler = new CommandHandler_ribbonButton7();

            RibbonButton ribbonButton8 = new RibbonButton();
            ribbonButton8.Id = "Создание дополнительных вершин в местах пересечения полилиний";
            ribbonButton8.Text = "Пересеч.линий";
            ribbonButton8.ShowText = true;
            ribbonButton8.Tag = ribbonButton8.Id;
            ribbonButton8.HelpTopic = "Создание дополнительных вершин в местах пересечения полилиний";
            // привязываем к кнопке обработчик нажатия
            ribbonButton8.CommandHandler = new CommandHandler_ribbonButton8();

            RibbonButton ribbonButton9 = new RibbonButton();
            ribbonButton9.Id = "Создание точек методом интерполяции";
            ribbonButton9.Text = "Интерпол+";
            ribbonButton9.ShowText = true;
            ribbonButton9.Tag = ribbonButton9.Id;
            ribbonButton9.HelpTopic = "Создание точек методом интерполяции";
            // привязываем к кнопке обработчик нажатия
            ribbonButton9.CommandHandler = new CommandHandler_ribbonButton9();

            RibbonButton ribbonButton10 = new RibbonButton();
            ribbonButton10.Id = "Групповое создание дополнительных вершин в местах пересечения полилиний";
            ribbonButton10.Text = "Пересеч.линий_Группа";
            ribbonButton10.ShowText = true;
            ribbonButton10.Tag = ribbonButton10.Id;
            ribbonButton10.HelpTopic = "Создание дополнительных вершин в местах пересечения полилиний";
            ribbonButton10.Description = "Создание дополнительных вершин в местах пересечения полилиний";
             // привязываем к кнопке обработчик нажатия
            ribbonButton10.CommandHandler = new CommandHandler_ribbonButton10();

            RibbonButton ribbonButton11 = new RibbonButton();
            ribbonButton11.Id = "Групповое создание точек в вершинах полилиний";
            ribbonButton11.Text = "Линия на ЦММ";
            ribbonButton11.ShowText = true;
            ribbonButton11.Tag = ribbonButton11.Id;
            ribbonButton11.HelpTopic = "Групповое создание точек в вершинах полилиний";
            ribbonButton11.Description = "Групповое создание точек в вершинах полилиний";
            // привязываем к кнопке обработчик нажатия
            ribbonButton11.CommandHandler = new CommandHandler_ribbonButton11();

            // создаем контейнер для элементов
            Autodesk.Windows.RibbonPanelSource rbPanelSource = new Autodesk.Windows.RibbonPanelSource();
            rbPanelSource.Title = "Создание поперечников";
            // добавляем в контейнер элементы управления
            rbPanelSource.Items.Add(ribbonButton5);
            rbPanelSource.Items.Add(new RibbonSeparator());
            rbPanelSource.Items.Add(ribbonButton6);
            rbPanelSource.Items.Add(new RibbonSeparator());
            rbPanelSource.Items.Add(ribbonButton7);
            rbPanelSource.Items.Add(new RibbonSeparator());
            rbPanelSource.Items.Add(ribbonButton1);
            rbPanelSource.Items.Add(new RibbonSeparator());
            rbPanelSource.Items.Add(ribbonButton2);
            rbPanelSource.Items.Add(new RibbonSeparator());
            rbPanelSource.Items.Add(ribbonButton4);
            rbPanelSource.Items.Add(new RibbonSeparator());
            rbPanelSource.Items.Add(ribbonButton3);
            rbPanelSource.Items.Add(new RibbonSeparator());
            


            // создаем контейнер для элементов
            Autodesk.Windows.RibbonPanelSource rbPanelSource_1 = new Autodesk.Windows.RibbonPanelSource();
            rbPanelSource_1.Title = "Создание линий сечения";
            // добавляем в контейнер элементы управления
            rbPanelSource_1.Items.Add(ribbonButton8);
            rbPanelSource_1.Items.Add(new RibbonSeparator());
            rbPanelSource_1.Items.Add(ribbonButton9);
            rbPanelSource_1.Items.Add(new RibbonSeparator());
            rbPanelSource_1.Items.Add(ribbonButton10);
            rbPanelSource_1.Items.Add(new RibbonSeparator());
            rbPanelSource_1.Items.Add(ribbonButton11);
            rbPanelSource_1.Items.Add(new RibbonSeparator());
            // создаем панель
            RibbonPanel rbPanel = new RibbonPanel();
            RibbonPanel rbPanel_1 = new RibbonPanel();
            // добавляем на панель контейнер для элементов
            rbPanel.Source = rbPanelSource;
            rbPanel_1.Source = rbPanelSource_1;

            // создаем вкладку
            RibbonTab rbTab = new RibbonTab();
            rbTab.Title = "Доп_функции";
            rbTab.Id = "AfoninRibbon";
            // добавляем на вкладку панель
            rbTab.Panels.Add(rbPanel);
            rbTab.Panels.Add(rbPanel_1);
            // получаем указатель на ленту AutoCAD
            RibbonControl rbCtrl = ComponentManager.Ribbon;
            // добавляем на ленту вкладку
            rbCtrl.Tabs.Add(rbTab);
            // делаем созданную вкладку активной ("выбранной")
            rbTab.IsActive = true;
#endif
        }
    }
}
