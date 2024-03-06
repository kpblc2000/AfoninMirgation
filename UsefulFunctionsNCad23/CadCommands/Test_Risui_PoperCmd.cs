using System;
using System.Globalization;
using System.IO;
using Infrastructure;

#if NCAD
using HostMgd.EditorInput;
using HostMgd.ApplicationServices;
using Teigha.DatabaseServices;
using Teigha.Geometry;
using Teigha.Runtime;
#elif ACAD
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
#endif

namespace UsefulFunctionsNCad23.CadCommands
{
    public static class Test_Risui_PoperCmd
    {
        [CommandMethod("Test_Risui_Poper", CommandFlags.Modal)]
        public static void Test_Risui_Poper()
        {
            //Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            //Editor ed = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor;
            //Database db = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            //   Editor ed= Teigha.Editor.CommandContext.Editor;
            Database db = Application.DocumentManager.MdiActiveDocument.Database;
            MessageService msgService = new MessageService();

            BlockTable acBlkTbl;   //объявляем переменные для базы с примитивами чертежа 
            BlockTableRecord acBlkTblRec;
            //-----далее нужно считать текстовый файл ***.pop и вычертить линию поперечника в зависимости от масштабов-----
            // Запрашиваем имя pop-файла

            PromptOpenFileOptions pfo = new PromptOpenFileOptions("\nВыберите файл поперечника pop-файл: \n");
            pfo.Filter = "pop-файлы (*.pop)|*.pop";
            //pfo.DialogName = "Выбор файла поперечника";
            pfo.DialogCaption = "Выбор файла поперечника";
            PromptFileNameResult pfr = ed.GetFileNameForOpen(pfo);

            if (pfr.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nНе удалось открыть выбранный файл\n");
                return;
            }
            //--------------- ed.WriteMessage("\nСчитан файл с названием: "+pfr.StringResult+"\n");
            // Читаем файл, получаем содержимое в виде массива строк
            string[] lines = File.ReadAllLines(pfr.StringResult);
            //---------------- ed.WriteMessage("\nПервая строка в считанном файле: \n" + lines[0]+ "\n");
            Form1 form1 = new Form1();
            //form1.Visible= true;
            //form1.groupBox2.Visible = true;
            if (Convert.ToDouble(form1.textBox1.Text) < 1 || Convert.ToDouble(form1.textBox1.Text) > 1000)
            {
                ed.WriteMessage("\nВ поле масштаба введено неверное значение");
                return;
            }
            double horizScaleKoef = 1000 / Convert.ToDouble(form1.textBox1.Text);
            if (Convert.ToDouble(form1.textBox2.Text) < 1 || Convert.ToDouble(form1.textBox2.Text) > 1000)
            {
                ed.WriteMessage("\nВ поле масштаба введено неверное значение");
                return;
            }
            double vertScaleKoef = 1000 / Convert.ToDouble(form1.textBox2.Text);


            using (Transaction Trans = db.TransactionManager.StartTransaction()) //чертеж поперечника создаем в рамках транзакции
            {
                acBlkTbl = (BlockTable)Trans.GetObject(db.BlockTableId, OpenMode.ForRead, false, true);      //открываем для чтения класс BlockTable
                acBlkTblRec = Trans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false, true) as BlockTableRecord;
                try
                {
                    //устанавливаем текущим стиль текста АТР, если он есть в базе чертежа
                    TextStyleTable textStyleTableDoc = (TextStyleTable)Trans.GetObject(db.TextStyleTableId, OpenMode.ForWrite, false, true);//открываем для чтения таблицу стилей текста
                    String sStyleName = "ATP";

                    if (textStyleTableDoc.Has(sStyleName))
                    {
                        db.Textstyle = textStyleTableDoc[sStyleName];
                    }
                    //_________________________________________________________________

                    Polyline profLine = new Polyline();//добавляем в базу полилинию, которая будет линией профиля
                    acBlkTblRec.AppendEntity(profLine);
                    Trans.AddNewlyCreatedDBObject(profLine, true);

                    string poperCaption = lines[0].Substring(5, lines[0].Length - 5);
                    ed.WriteMessage($"\nСчитано значение заголовка поперечника: {poperCaption}\n");
                    //задаем стартовую точку для создания линии поперечника_____________________________________________
                    PromptPointOptions startPointOpt = new PromptPointOptions("\nУкажите стартовую точку 0 поперечника\n");
                    startPointOpt.AllowNone = false;
                    PromptPointResult startPointRes = ed.GetPoint(startPointOpt);
                    if (startPointRes.Status != PromptStatus.OK) return;
                    Point3d startPoint3D = startPointRes.Value;
                    //______________________Считываем из выбранного пользователем файла *.рор и раскидываем
                    //их по соответствующим коллекциям с ObjectId текстов, строя основную линию профиля__________________________________
                    //double minY= startPoint3D.Y + Convert.ToDouble(znach[2]) * vertScaleKoef;
                    double minY = double.MaxValue;
                    double maxY = double.MinValue;
                    bool changeZnak = true;
                    //DBText[] dBTextsOpisanie = new DBText[] { };
                    ObjectIdCollection colIdOpisanie = new ObjectIdCollection();
                    // DBText[] dBTextsOtmetki = new DBText[] { };
                    ObjectIdCollection colIdOtmetki = new ObjectIdCollection();
                    // DBText[] dBTextsRasst = new DBText[] { };
                    ObjectIdCollection colIdRasst = new ObjectIdCollection();
                    //DBText[] dBTextsCodes = new DBText[] { };
                    ObjectIdCollection colIdCodes = new ObjectIdCollection();
                    ObjectIdCollection colIdPodzemkaRasst = new ObjectIdCollection();
                    ObjectIdCollection colIdPodzemkaAbsolut = new ObjectIdCollection();
                    ObjectIdCollection colIdRasstIsso = new ObjectIdCollection();

                    ed.WriteMessage($"Файл с поперечником {poperCaption} содержит {lines.Length} строк\n");

                    for (int i = 1; i < lines.Length; i++)
                    {

                        // char[] charsToTrim = { '*', ' ', '\'' };
                        // char[] charsToTrim = {' '};
                        lines[i].Trim(new char[] { ' ' });

                        // ed.WriteMessage($"Считывается строка с индексом {i}\n");

                        if (lines[i][0] == '1' || lines[i][0] == '0' || lines[i][0] == '2' || lines[i][0] == '3' || lines[i][0] == '4' || lines[i][0] == '5' || lines[i][0] == '6' || lines[i][0] == '7' || lines[i][0] == '8' || lines[i][0] == '9')//если строка начинается с символа 1 или 0 (то есть код есть), то разбиваем массив на подстроки
                        {
                            string[] znach;
                            znach = lines[i].Split(new char[] { ' ' }, 4, StringSplitOptions.RemoveEmptyEntries);//разбиваем строку на подстроки по разделитель "пробел",
                                                                                                                 //максимум 4 подстроки (код,расстояние, высота, описание),
                                                                                                                 //последовательные пробелы исключаются

                            if (znach.Length > 1)
                            {


                                double Code = Convert.ToDouble(znach[0]);
                                double deltaX = Math.Round(Convert.ToDouble(znach[1]), 2);//меняем знак расстояния в зависимости от "лево" или "право"
                                if (deltaX == 0)
                                {
                                    changeZnak = false;
                                }
                                if (changeZnak == true)
                                {
                                    deltaX = -deltaX;
                                }
                                double deltaY = Math.Round(Convert.ToDouble(znach[2]), 2);
                                if (Code == 0 || Code == 1 || Code == 501)
                                {
                                    minY = Math.Min(minY, startPoint3D.Y + deltaY * vertScaleKoef);
                                    maxY = Math.Max(maxY, startPoint3D.Y + deltaY * vertScaleKoef);
                                }

                                string opisanie = "";
                                if (znach.Length > 3)
                                {
                                    opisanie = znach[3];
                                }
                                //ed.WriteMessage($"\nСчитано следующее содержимое поперечника:\n {Code}  {deltaX}  {deltaY} {opisanie}\n");

                                //в случае, если код точки - не подземка и не ИССО, то в линию поперечника добавляем вершину_________________

                                if (Code == 0 || Code == 2 || Code == 3 || Code == 4 || Code == 5 || Code == 6 || Code == 7 || Code == 8 || Code == 9 || Code == 10 || Code == 11 || Code == 15)
                                {

                                    int indVert = 0;
                                    Point2d poperVertex = new Point2d(startPoint3D.X + deltaX * horizScaleKoef, startPoint3D.Y + deltaY * vertScaleKoef);
                                    profLine.AddVertexAt(indVert, poperVertex, 0, 0, 0);
                                    indVert++;
                                }
                                //_______________вставляем текст с описанием точки, если есть__________________________
                                if (opisanie != "")
                                {
                                    DBText newOpisanie = new DBText();
                                    acBlkTblRec.AppendEntity(newOpisanie);
                                    Trans.AddNewlyCreatedDBObject(newOpisanie, true);
                                    newOpisanie.TextString = opisanie;
                                    if (Code == 120 || Code == 121 || Code == 130 || Code == 131 || Code == 140 || Code == 141)
                                    {
                                        String roundDist = Math.Abs(deltaX).ToString("0.00", CultureInfo.InvariantCulture);
                                        newOpisanie.TextString = roundDist + " " + newOpisanie.TextString;
                                        // newOpisanie.TextString = Convert.ToString(Abs(deltaX)) + " " + newOpisanie.TextString;
                                    }
                                    else if (Code == 1)
                                    {
                                        newOpisanie.TextString = "путь " + newOpisanie.TextString;
                                    }
                                    // newOpisanie.Height = 2;
                                    newOpisanie.Height = Convert.ToDouble(form1.textBox3.Text);
                                    newOpisanie.Rotation = 1.5708;
                                    newOpisanie.Position = new Point3d(startPoint3D.X + deltaX * horizScaleKoef, startPoint3D.Y + deltaY * vertScaleKoef, 0);
                                    //dBTextsOpisanie.Append(newOpisanie);
                                    if (newOpisanie.TextStyleName == "ATP")
                                    {
                                        newOpisanie.Oblique = 0.2618;
                                        newOpisanie.WidthFactor = 0.8;
                                    }

                                    colIdOpisanie.Add(newOpisanie.ObjectId);


                                }
                                //_______________вставляем текст с отметкой точки, если больше 0__________________________
                                if (Code == 0 || Code == 1 || Code == 2 || Code == 3 || Code == 4 || Code == 5 || Code == 6 || Code == 7 || Code == 8 || Code == 9 || Code == 10 || Code == 11 || Code == 15) //здесь был код 501
                                {
                                    DBText newOtmetka = new DBText();
                                    acBlkTblRec.AppendEntity(newOtmetka);
                                    Trans.AddNewlyCreatedDBObject(newOtmetka, true);
                                    newOtmetka.TextString = Convert.ToString(Math.Round(deltaY, 2));
                                    newOtmetka.TextString = deltaY.ToString("0.00", CultureInfo.InvariantCulture);
                                    if (Code == 1)
                                    {
                                        newOtmetka.TextString = "СГР " + newOtmetka.TextString;
                                    }
                                    newOtmetka.Height = Convert.ToDouble(form1.textBox3.Text);
                                    newOtmetka.Rotation = 1.5708;
                                    newOtmetka.ColorIndex = 1;
                                    newOtmetka.Position = new Point3d(startPoint3D.X + deltaX * horizScaleKoef, startPoint3D.Y - 50 + deltaY * vertScaleKoef, 0);
                                    //dBTextsOtmetki.Append(newOtmetka);
                                    if (newOtmetka.TextStyleName == "ATP")
                                    {
                                        newOtmetka.Oblique = 0.2618;
                                        newOtmetka.WidthFactor = 0.8;
                                    }
                                    colIdOtmetki.Add(newOtmetka.ObjectId);

                                }
                                if (Code == 120 || Code == 121 || Code == 130 || Code == 131 || Code == 140 || Code == 141 || Code == 501) //делаем отдельную коллекцию для превышений/заглублений коммуникаций
                                {
                                    DBText newOtmetka = new DBText();
                                    acBlkTblRec.AppendEntity(newOtmetka);
                                    Trans.AddNewlyCreatedDBObject(newOtmetka, true);
                                    newOtmetka.TextString = Convert.ToString(Math.Round(deltaY, 2));
                                    newOtmetka.TextString = deltaY.ToString("0.00", CultureInfo.InvariantCulture);
                                    newOtmetka.Height = Convert.ToDouble(form1.textBox3.Text);
                                    newOtmetka.Rotation = 1.5708;
                                    newOtmetka.ColorIndex = 102;
                                    if (newOtmetka.TextStyleName == "ATP")
                                    {
                                        newOtmetka.Oblique = 0.2618;
                                        newOtmetka.WidthFactor = 0.8;
                                    }
                                    newOtmetka.Position = new Point3d(startPoint3D.X + deltaX * horizScaleKoef, startPoint3D.Y - 50 + deltaY * vertScaleKoef, 0);
                                    //dBTextsOtmetki.Append(newOtmetka);
                                    colIdPodzemkaRasst.Add(newOtmetka.ObjectId);
                                    if (Code == 121 || Code == 131 || Code == 141 || Code == 501)
                                    {
                                        Point3d absPodzemkaPt = new Point3d(startPoint3D.X + deltaX * horizScaleKoef, startPoint3D.Y + deltaY * vertScaleKoef, 0);
                                        DBPoint absPodzemkaDBp = new DBPoint(absPodzemkaPt);
                                        acBlkTblRec.AppendEntity(absPodzemkaDBp);
                                        Trans.AddNewlyCreatedDBObject(absPodzemkaDBp, true);
                                        colIdPodzemkaAbsolut.Add(absPodzemkaDBp.ObjectId);
                                    }


                                }
                                //_______________вставляем текст с расстоянием до 0-точки (далее мы их переделаем в расстояния между точками перелома)
                                // если код предусматривает другое - пропускаем__________________________
                                if ((Code != 120) && (Code != 121) && (Code != 130) && (Code != 131) && (Code != 140) && (Code != 141) && (Code != 501))
                                {
                                    DBText newRasst = new DBText();
                                    acBlkTblRec.AppendEntity(newRasst);
                                    Trans.AddNewlyCreatedDBObject(newRasst, true);
                                    newRasst.TextString = Convert.ToString(Math.Round(deltaX, 2));

                                    newRasst.Height = Convert.ToDouble(form1.textBox3.Text);
                                    newRasst.Rotation = 1.5708;
                                    newRasst.ColorIndex = 2;
                                    newRasst.Position = new Point3d(startPoint3D.X + deltaX * horizScaleKoef, startPoint3D.Y - 60 + deltaY * vertScaleKoef, 0);
                                    //dBTextsRasst.Append(newRasst);
                                    //int indOtmMas=dBTextsOtmetki.GetUpperBound(0);
                                    //ed.WriteMessage($"\nСейчас верхний индекс массива с расстояниями: {indOtmMas}\n");
                                    colIdRasst.Add(newRasst.ObjectId);
                                }
                                if (Code == 501) //в отдельную коллекцию собираем расстояния до точек с кодом ИССО. 
                                {
                                    DBText newRasst = new DBText();
                                    acBlkTblRec.AppendEntity(newRasst);
                                    Trans.AddNewlyCreatedDBObject(newRasst, true);
                                    newRasst.TextString = Convert.ToString(Math.Round(deltaX, 2));

                                    newRasst.Height = Convert.ToDouble(form1.textBox3.Text);
                                    newRasst.Rotation = 1.5708;
                                    newRasst.ColorIndex = 2;
                                    newRasst.Position = new Point3d(startPoint3D.X + deltaX * horizScaleKoef, startPoint3D.Y - 60 + deltaY * vertScaleKoef, 0);
                                    //dBTextsRasst.Append(newRasst);
                                    //int indOtmMas=dBTextsOtmetki.GetUpperBound(0);
                                    //ed.WriteMessage($"\nСейчас верхний индекс массива с расстояниями: {indOtmMas}\n");
                                    colIdRasstIsso.Add(newRasst.ObjectId);

                                }

                                // сохраняем данные о кодах в массив текстов__________________________

                                DBText newCode = new DBText();
                                acBlkTblRec.AppendEntity(newCode);
                                Trans.AddNewlyCreatedDBObject(newCode, true);
                                newCode.TextString = Convert.ToString(Code);
                                newCode.Height = Convert.ToDouble(form1.textBox3.Text);
                                newCode.Rotation = -1.5708;
                                newCode.ColorIndex = 3;

                                newCode.Position = new Point3d(startPoint3D.X + deltaX * horizScaleKoef, startPoint3D.Y - 70 + deltaY * vertScaleKoef, 0);
                                // dBTextsCodes.Append(newCode);
                                colIdCodes.Add(newCode.ObjectId);
                                // ed.WriteMessage($"Считаны следующие значения: Код: {Code};\t расстояние: {deltaX};\tОтметка: {deltaY};\t Пояснение: {opisanie}\n");
                            }
                        }




                    }

                    //ed.WriteMessage($"Вычислено значение переменной MinY={minY}\n");
                    // ed.WriteMessage($"Установлено значение MinY: {minY};\n начальная точки по Х: {profLine.StartPoint.X}\n");
                    ChertiPoper(profLine, minY, maxY, colIdOpisanie, colIdOtmetki, colIdRasst, colIdCodes, Trans, acBlkTblRec, horizScaleKoef, vertScaleKoef, colIdPodzemkaRasst, colIdPodzemkaAbsolut, colIdRasstIsso, db.Textstyle, poperCaption, startPoint3D, msgService);
                    // ed.WriteMessage($"\nДлина коллекции с описаниями ={colIdOpisanie.Count}\n, длина массива с отметками = {colIdOtmetki.Count}\n, длина массива с расстояниями {colIdRasst.Count}\n, длина массива с кодами {colIdCodes.Count}\n");
                    Trans.Commit();
                    // docklock.Dispose();
                }
                catch (System.Exception ex1)
                {
                    msgService.ExceptionMessage(ex1, $"В процессе считывания данных поперечника возникла ошибка");
                    Trans.Abort();
                }
            }



        }

        private static void ChertiPoper(Polyline poperLine, double minY, double maxY, ObjectIdCollection colIdOpisanie, ObjectIdCollection colIdOtmetki, ObjectIdCollection colIdRasst, ObjectIdCollection colIdCodes, Transaction Trans, BlockTableRecord acBlkTblRec, double horizScaleKoef, double vertScaleKoef, ObjectIdCollection colIdPodzemkaRasst, ObjectIdCollection colIdPodzemkaAbsolut, ObjectIdCollection colIdRasstIsso, ObjectId CurTextstyle, string poperCaption, Point3d startPoint3D, MessageService msgService)
        {
            ObjectIdCollection shrixiPutey = new ObjectIdCollection();
            ObjectIdCollection txtCodePut = new ObjectIdCollection();
            ObjectIdCollection pointsFinalPunktir = new ObjectIdCollection();
            ObjectIdCollection mezhputTEXTs = new ObjectIdCollection();
            ObjectIdCollection rasstTEXTs = new ObjectIdCollection();
            ObjectIdCollection allPoperEnts = new ObjectIdCollection();
            //начинаем с создания линии горизонта профиля - на 60мм ниже минимальной координаты Y______________________________

            Line gorizontLine = new Line(new Point3d(poperLine.StartPoint.X, minY - 60, 0), new Point3d(poperLine.EndPoint.X, minY - 60, 0));
            acBlkTblRec.AppendEntity(gorizontLine);
            Trans.AddNewlyCreatedDBObject(gorizontLine, true);
            allPoperEnts.Add(gorizontLine.ObjectId);
            //_____________________________________________________________________________________________________________________
            //----чертим вспомогательные линии---
            Line vspomLine1 = new Line(new Point3d(gorizontLine.StartPoint.X, gorizontLine.StartPoint.Y - 15, 0), new Point3d(gorizontLine.EndPoint.X, gorizontLine.StartPoint.Y - 15, 0));
            acBlkTblRec.AppendEntity(vspomLine1);
            Trans.AddNewlyCreatedDBObject(vspomLine1, true);
            allPoperEnts.Add(vspomLine1.ObjectId);
            Line vspomLine2 = new Line(new Point3d(gorizontLine.StartPoint.X, gorizontLine.StartPoint.Y - 20, 0), new Point3d(gorizontLine.EndPoint.X, gorizontLine.StartPoint.Y - 20, 0));
            acBlkTblRec.AppendEntity(vspomLine2);
            Trans.AddNewlyCreatedDBObject(vspomLine2, true);
            allPoperEnts.Add(vspomLine2.ObjectId);
            Line vspomLine3 = new Line(new Point3d(gorizontLine.StartPoint.X, gorizontLine.StartPoint.Y - 35, 0), new Point3d(gorizontLine.EndPoint.X, gorizontLine.StartPoint.Y - 35, 0));
            acBlkTblRec.AppendEntity(vspomLine3);
            Trans.AddNewlyCreatedDBObject(vspomLine3, true);
            allPoperEnts.Add(vspomLine3.ObjectId);
            Line vspomLine4 = new Line(new Point3d(gorizontLine.StartPoint.X, gorizontLine.StartPoint.Y - 40, 0), new Point3d(gorizontLine.EndPoint.X, gorizontLine.StartPoint.Y - 40, 0));
            acBlkTblRec.AppendEntity(vspomLine4);
            Trans.AddNewlyCreatedDBObject(vspomLine4, true);
            allPoperEnts.Add(vspomLine4.ObjectId);

            allPoperEnts.Add(poperLine.ObjectId);

            Form1 form1 = new Form1();

            //ed.WriteMessage("Закончили рисовать линии\n");
            //_________________ставим на место все тексты с описаниями (на 3 выше горизонта)____________________________
            foreach (ObjectId opisId in colIdOpisanie)
            {
                DBText myOpisTxt = (DBText)Trans.GetObject(opisId, OpenMode.ForWrite);
                // ed.WriteMessage($"Считан текст с содержимым {myOpisTxt.TextString}\n");
                myOpisTxt.Position = new Point3d(myOpisTxt.Position.X, gorizontLine.StartPoint.Y + 3, 0);
                allPoperEnts.Add(opisId);
            }

            foreach (ObjectId podzemTXTs in colIdPodzemkaRasst)
            {
                DBText myPodzemTxt = (DBText)Trans.GetObject(podzemTXTs, OpenMode.ForWrite);
                // ed.WriteMessage($"Считан текст с содержимым {myOpisTxt.TextString}\n");
                myPodzemTxt.Position = new Point3d(myPodzemTxt.Position.X, maxY + 200, 0);
                if (myPodzemTxt.TextStyleName == "ATP")
                {
                    myPodzemTxt.Oblique = 0.2618;
                    myPodzemTxt.WidthFactor = 0.8;
                }
            }
            foreach (ObjectId RasstIssoTXTs in colIdRasstIsso)
            {
                DBText myRasstIssoTxt = (DBText)Trans.GetObject(RasstIssoTXTs, OpenMode.ForWrite);
                // ed.WriteMessage($"Считан текст с содержимым {myOpisTxt.TextString}\n");
                myRasstIssoTxt.Position = new Point3d(myRasstIssoTxt.Position.X, maxY + 150, 0);
                if (myRasstIssoTxt.TextStyleName == "ATP")
                {
                    myRasstIssoTxt.Oblique = 0.2618;
                    myRasstIssoTxt.WidthFactor = 0.8;
                }
            }

            // ed.WriteMessage("Закончили колупать описания\n");
            //_________________ставим на место все тексты с отметками (на 33 ниже горизонта)____________________________
            foreach (ObjectId otmetId in colIdOtmetki)
            {
                DBText myOtmetTxt = (DBText)Trans.GetObject(otmetId, OpenMode.ForWrite);
                myOtmetTxt.Position = new Point3d(myOtmetTxt.Position.X, gorizontLine.StartPoint.Y - 33, 0);
                myOtmetTxt.ColorIndex = 256;
                //далее передвигаем значение на середину, чтобы текст визуально был по центру ячейки
                Point3d leftPt = myOtmetTxt.GeometricExtents.MinPoint;
                Point3d rightPt = myOtmetTxt.GeometricExtents.MaxPoint;
                double deltaX = 0.5 * (rightPt.X - leftPt.X);
                allPoperEnts.Add(otmetId);
                myOtmetTxt.Position = new Point3d(myOtmetTxt.Position.X + deltaX, myOtmetTxt.Position.Y, myOtmetTxt.Position.Z);
                if (myOtmetTxt.TextString.Substring(0, 3) == "СГР")
                {
                    myOtmetTxt.Position = new Point3d(myOtmetTxt.GeometricExtents.MinPoint.X, gorizontLine.StartPoint.Y + 3, 0);
                }

                if (myOtmetTxt.TextStyleName == "ATP")
                {
                    myOtmetTxt.Oblique = 0.2618;
                    myOtmetTxt.WidthFactor = 0.8;
                }

                // ed.WriteMessage($"Изменена позиция текста с отметкой на {myOtmetTxt.Position}\n");
            }

            //-----------создаем тексты с расстояниями---------------------------
            for (int i = 0; i < colIdRasst.Count - 1; i++)
            {
                DBText myDistTxt1 = (DBText)Trans.GetObject(colIdRasst[i], OpenMode.ForWrite);
                DBText myDistTxt2 = (DBText)Trans.GetObject(colIdRasst[i + 1], OpenMode.ForWrite);
                double isxD1 = Math.Round(Convert.ToDouble(myDistTxt1.TextString), 2);
                double isxD2 = Math.Round(Convert.ToDouble(myDistTxt2.TextString), 2);
                double myDist = Math.Round(Math.Abs(isxD1 - isxD2), 2);
                if (myDist > 0)//-------делаем так, чтобы не было расстояний 0.00------------------------------------------------
                {
                    DBText myDistTxt = new DBText();
                    myDistTxt.Rotation = 0;
                    myDistTxt.ColorIndex = 200;
                    myDistTxt.Height = Convert.ToDouble(form1.textBox3.Text);
                    myDistTxt.TextString = myDist.ToString("0.00", CultureInfo.InvariantCulture);
                    myDistTxt.Position = new Point3d(myDistTxt1.Position.X + 0.5 * (myDistTxt2.Position.X - myDistTxt1.Position.X), gorizontLine.StartPoint.Y - 39, 0);
                    acBlkTblRec.AppendEntity(myDistTxt);
                    Trans.AddNewlyCreatedDBObject(myDistTxt, true);
                    //далее передвигаем значение на середину, чтобы текст визуально был по центру ячейки
                    Point3d leftPt = myDistTxt.GeometricExtents.MinPoint;
                    Point3d rightPt = myDistTxt.GeometricExtents.MaxPoint;
                    double deltaX = 0.5 * (rightPt.X - leftPt.X);
                    myDistTxt.Position = new Point3d(myDistTxt.Position.X - deltaX, myDistTxt.Position.Y, myDistTxt.Position.Z);
                    rasstTEXTs.Add(myDistTxt.ObjectId);
                    allPoperEnts.Add(myDistTxt.ObjectId);
                    if (myDistTxt.TextStyleName == "ATP")
                    {
                        myDistTxt.Oblique = 0.2618;
                        myDistTxt.WidthFactor = 0.8;
                    }
                }

            }
            foreach (ObjectId rasstID in colIdRasst)
            {
                DBText myRasstTxt = (DBText)Trans.GetObject(rasstID, OpenMode.ForWrite);
                myRasstTxt.ColorIndex = 180;
                if (myRasstTxt.TextStyleName == "ATP")
                {
                    myRasstTxt.Oblique = 0.2618;
                    myRasstTxt.WidthFactor = 0.8;
                }
            }


            //___________________________________________________________________________________________________________________
            foreach (ObjectId codeId in colIdCodes)
            {
                DBText myCodeTxt = (DBText)Trans.GetObject(codeId, OpenMode.ForWrite);
                // ed.WriteMessage($"Считан текст с содержимым {myOpisTxt.TextString}\n");
                myCodeTxt.Position = new Point3d(myCodeTxt.Position.X, maxY + 300, 0);

                //------------Создание вертикальных палок-------------------
                if (myCodeTxt.TextString == "0" || myCodeTxt.TextString == "2" || myCodeTxt.TextString == "3" || myCodeTxt.TextString == "4" || myCodeTxt.TextString == "5" || myCodeTxt.TextString == "6" || myCodeTxt.TextString == "7" || myCodeTxt.TextString == "8" || myCodeTxt.TextString == "9" || myCodeTxt.TextString == "10" || myCodeTxt.TextString == "11" || myCodeTxt.TextString == "15")
                {
                    using (Line AllVertLine = new Line(myCodeTxt.Position, new Point3d(myCodeTxt.Position.X, vspomLine4.StartPoint.Y, 0)))
                    {
                        //сначала делаем основные палки сверху от самой линии поперечника до первого горизонта
                        Point3dCollection intCollection = new Point3dCollection();
                        AllVertLine.IntersectWith(poperLine, intersectType: Intersect.OnBothOperands, intCollection, IntPtr.Zero, IntPtr.Zero);
                        if (intCollection.Count > 0)
                        {
                            Point3d point3Dfirst = intCollection[0]; //проблема!
                            Point3d point3Dsecond = new Point3d(myCodeTxt.Position.X, gorizontLine.StartPoint.Y, 0);
                            Line line1 = new Line(point3Dfirst, point3Dsecond);
                            acBlkTblRec.AppendEntity(line1);
                            Trans.AddNewlyCreatedDBObject(line1, true);
                            allPoperEnts.Add(line1.ObjectId);
                        }
                        else
                        {
                            msgService.ConsoleMessage($"Не нашлось пересечения для точки с кодом {myCodeTxt.TextString} \n");
                            // break;
                        }

                        //теперь делаем коротенькие палочки между 3 и 4 линиями
                        Line line2 = new Line(new Point3d(myCodeTxt.Position.X, vspomLine3.StartPoint.Y, 0), new Point3d(myCodeTxt.Position.X, vspomLine4.StartPoint.Y, 0));
                        acBlkTblRec.AppendEntity(line2);
                        Trans.AddNewlyCreatedDBObject(line2, true);
                        allPoperEnts.Add(line2.ObjectId);
                    }
                }

                //---------создаем пунктиры ж-д путей и маленькие линии для кода 1 между 3 и 4 линиями-----------
                if (myCodeTxt.TextString == "1")
                {
                    //теперь делаем коротенькие палочки между 3 и 4 линиями
                    Line line2 = new Line(new Point3d(myCodeTxt.Position.X, vspomLine3.StartPoint.Y, 0), new Point3d(myCodeTxt.Position.X, vspomLine4.StartPoint.Y, 0));
                    acBlkTblRec.AppendEntity(line2);
                    Trans.AddNewlyCreatedDBObject(line2, true);
                    allPoperEnts.Add(line2.ObjectId);
                    //теперь надо сделать пунктир как серию коротких штрихов от линии горизонта до верха (maxY+3м в масштабе)
                    //закидываем все создаваемые штрихи в отдельную коллекцию на будущее

                    double startPunktirY = gorizontLine.StartPoint.Y;
                    while (startPunktirY < (maxY + 15))
                    {
                        Point3d point3DStartPunk = new Point3d(myCodeTxt.Position.X, startPunktirY, 0);
                        //point3DStartPunk = new Point3d(myCodeTxt.Position.X, gorizontLine.StartPoint.Y, 0);
                        Point3d point2 = new Point3d(point3DStartPunk.X, point3DStartPunk.Y + 3, 0);
                        Line shtrix = new Line(point3DStartPunk, point2);
                        acBlkTblRec.AppendEntity(shtrix);
                        Trans.AddNewlyCreatedDBObject(shtrix, true);
                        shrixiPutey.Add(shtrix.ObjectId);
                        allPoperEnts.Add(shtrix.ObjectId);
                        startPunktirY += 4;
                        if (startPunktirY >= maxY + 15)
                        {
                            DBPoint pointFinalPunktir = new DBPoint(point2);
                            acBlkTblRec.AppendEntity(pointFinalPunktir);
                            Trans.AddNewlyCreatedDBObject(pointFinalPunktir, true);
                            pointsFinalPunktir.Add(pointFinalPunktir.ObjectId);
                        }
                    }
                    //--------------добавляем каждый текст с кодом 1 в отдельную коллекцию, чтобы сверху нарисовать междопутье-----------------------------------
                    txtCodePut.Add(myCodeTxt.ObjectId);
                }
                //----для следующего кода делаем условный знак (подземка, ЛЭП и прочее)
                //-----рисуем обозначение кода 120 (подземная коммуникация с заглублением)----------------------
                if (myCodeTxt.TextString == "120" || myCodeTxt.TextString == "121" || myCodeTxt.TextString == "130" || myCodeTxt.TextString == "131" || myCodeTxt.TextString == "140" || myCodeTxt.TextString == "141")
                {
                    double deltaH = 0;
                    //ищем в коллекции colIdPodzemkaRasst соответствующее коду значение заглубления
                    for (int i = 0; i < colIdPodzemkaRasst.Count; i++)
                    {
                        DBText myPodzemTxt = (DBText)Trans.GetObject(colIdPodzemkaRasst[i], OpenMode.ForWrite);
                        if (myPodzemTxt.Position.X == myCodeTxt.Position.X)
                        {
                            deltaH = Convert.ToDouble(myPodzemTxt.TextString);
                            break;
                        }
                    }

                    //теперь получив заглубление/превышение для каждого значения кода, в зависимости от конкретного кода делаем условный знак

                    if (myCodeTxt.TextString == "120")
                    {
                        using (Line AllVertLine = new Line(myCodeTxt.Position, new Point3d(myCodeTxt.Position.X, vspomLine4.StartPoint.Y, 0)))
                        {
                            //сначала находим вершину заглубления коммуникации
                            Point3dCollection intCollection = new Point3dCollection();
                            AllVertLine.IntersectWith(poperLine, intersectType: Intersect.OnBothOperands, intCollection, IntPtr.Zero, IntPtr.Zero);
                            Point3d point3D_INTER = intCollection[0]; //проблема! (при чтении традиционных рор-файлов
                            Point3d point3D_Down = new Point3d(myCodeTxt.Position.X, gorizontLine.StartPoint.Y, 0);
                            Point3d point3D_Up = new Point3d(point3D_INTER.X, point3D_INTER.Y + vertScaleKoef * deltaH - 0.3, 0);
                            Circle roundCir = new Circle(point3D_Up, Vector3d.ZAxis, 0.3);
                            acBlkTblRec.AppendEntity(roundCir);
                            Trans.AddNewlyCreatedDBObject(roundCir, true);
                            allPoperEnts.Add(roundCir.ObjectId);

                            Line line1 = new Line(point3D_Down, new Point3d(point3D_Up.X, point3D_Up.Y - 0.3, 0));
                            acBlkTblRec.AppendEntity(line1);
                            Trans.AddNewlyCreatedDBObject(line1, true);
                            allPoperEnts.Add(line1.ObjectId);
                        }

                    }
                    if (myCodeTxt.TextString == "121")//для случая, когда задана абсолютная отметка верха коммуникации,
                                                      //строим условный знак до точки из коллекции colIdPodzemkaAbsolut
                    {
                        //ищем в коллекции colIdPodzemkaAbsolut соответствующую точку
                        for (int i = 0; i < colIdPodzemkaAbsolut.Count; i++)
                        {
                            DBPoint myPodzemPt = (DBPoint)Trans.GetObject(colIdPodzemkaAbsolut[i], OpenMode.ForWrite);
                            if (myPodzemPt.Position.X == myCodeTxt.Position.X)
                            {
                                Point3d point3D_Up = new Point3d(myPodzemPt.Position.X, myPodzemPt.Position.Y - 0.3, 0);
                                Circle roundCir = new Circle(point3D_Up, Vector3d.ZAxis, 0.3);
                                acBlkTblRec.AppendEntity(roundCir);
                                Trans.AddNewlyCreatedDBObject(roundCir, true);
                                allPoperEnts.Add(roundCir.ObjectId);
                                Point3d point3D_Down = new Point3d(myCodeTxt.Position.X, gorizontLine.StartPoint.Y, 0);
                                Line line1 = new Line(point3D_Down, new Point3d(point3D_Up.X, point3D_Up.Y - 0.3, 0));
                                acBlkTblRec.AppendEntity(line1);
                                Trans.AddNewlyCreatedDBObject(line1, true);
                                allPoperEnts.Add(line1.ObjectId);
                                break;
                            }
                        }
                    }
                    //далее строим код 130 и 131
                    if (myCodeTxt.TextString == "130")
                    {
                        using (Line AllVertLine = new Line(myCodeTxt.Position, new Point3d(myCodeTxt.Position.X, vspomLine4.StartPoint.Y, 0)))
                        {
                            //сначала находим вершину верха коммуникации
                            Point3dCollection intCollection = new Point3dCollection();
                            AllVertLine.IntersectWith(poperLine, intersectType: Intersect.OnBothOperands, intCollection, IntPtr.Zero, IntPtr.Zero);
                            Point3d point3D_INTER = intCollection[0]; //проблема! (при чтении традиционных рор-файлов
                            Point3d point3D_Down = new Point3d(myCodeTxt.Position.X, gorizontLine.StartPoint.Y, 0);
                            Point3d point3D_Up = new Point3d(point3D_INTER.X, point3D_INTER.Y + vertScaleKoef * deltaH, 0);
                            Line lineLep = new Line(point3D_Down, point3D_Up);//строится вертикальная линия от горизонта до верхней точки с превышением
                            acBlkTblRec.AppendEntity(lineLep);
                            Trans.AddNewlyCreatedDBObject(lineLep, true);
                            allPoperEnts.Add(lineLep.ObjectId);

                            Point3d point3DLeftLep = new Point3d(point3D_Up.X - 5, point3D_Up.Y, 0);
                            Point3d point3DRightLep = new Point3d(point3D_Up.X + 5, point3D_Up.Y, 0);
                            Line lineLepGor = new Line(point3DLeftLep, point3DRightLep);//горизонтальная линия значка ЛЭП
                            acBlkTblRec.AppendEntity(lineLepGor);
                            Trans.AddNewlyCreatedDBObject(lineLepGor, true);
                            allPoperEnts.Add(lineLepGor.ObjectId);

                            Point3d point3DleftCir = new Point3d(point3D_Up.X - 4, point3D_Up.Y - 1, 0);
                            Point3d point3DRightCir = new Point3d(point3D_Up.X + 4, point3D_Up.Y - 1, 0);
                            Circle roundCirLeft = new Circle(point3DleftCir, Vector3d.ZAxis, 0.5);
                            Circle roundCirRight = new Circle(point3DRightCir, Vector3d.ZAxis, 0.5);//слева и справа рисуем окружности значка ЛЭП
                            acBlkTblRec.AppendEntity(roundCirLeft);
                            Trans.AddNewlyCreatedDBObject(roundCirLeft, true);
                            allPoperEnts.Add(roundCirLeft.ObjectId);

                            acBlkTblRec.AppendEntity(roundCirRight);
                            Trans.AddNewlyCreatedDBObject(roundCirRight, true);
                            allPoperEnts.Add(roundCirRight.ObjectId);
                        }

                    }
                    if (myCodeTxt.TextString == "131")//для случая, когда задана абсолютная отметка верха коммуникации,
                                                      //строим условный знак до точки из коллекции colIdPodzemkaAbsolut
                    {
                        //ищем в коллекции colIdPodzemkaAbsolut соответствующую точку
                        for (int i = 0; i < colIdPodzemkaAbsolut.Count; i++)
                        {
                            DBPoint myPodzemPt = (DBPoint)Trans.GetObject(colIdPodzemkaAbsolut[i], OpenMode.ForWrite);
                            if (myPodzemPt.Position.X == myCodeTxt.Position.X)
                            {
                                Point3d point3D_Up = new Point3d(myPodzemPt.Position.X, myPodzemPt.Position.Y, 0);//найдена верхняя абсолютная точка
                                Point3d point3D_Down = new Point3d(myCodeTxt.Position.X, gorizontLine.StartPoint.Y, 0);
                                Line lineLep = new Line(point3D_Down, point3D_Up);//строится вертикальная линия от горизонта до верхней точки с превышением
                                acBlkTblRec.AppendEntity(lineLep);
                                Trans.AddNewlyCreatedDBObject(lineLep, true);
                                allPoperEnts.Add(lineLep.ObjectId);

                                Point3d point3DLeftLep = new Point3d(point3D_Up.X - 5, point3D_Up.Y, 0);
                                Point3d point3DRightLep = new Point3d(point3D_Up.X + 5, point3D_Up.Y, 0);
                                Line lineLepGor = new Line(point3DLeftLep, point3DRightLep);//горизонтальная линия значка ЛЭП
                                acBlkTblRec.AppendEntity(lineLepGor);
                                Trans.AddNewlyCreatedDBObject(lineLepGor, true);
                                allPoperEnts.Add(lineLepGor.ObjectId);

                                Point3d point3DleftCir = new Point3d(point3D_Up.X - 4, point3D_Up.Y - 1, 0);
                                Point3d point3DRightCir = new Point3d(point3D_Up.X + 4, point3D_Up.Y - 1, 0);
                                Circle roundCirLeft = new Circle(point3DleftCir, Vector3d.ZAxis, 0.5);
                                Circle roundCirRight = new Circle(point3DRightCir, Vector3d.ZAxis, 0.5);//слева и справа рисуем окружности значка ЛЭП
                                acBlkTblRec.AppendEntity(roundCirLeft);
                                Trans.AddNewlyCreatedDBObject(roundCirLeft, true);
                                allPoperEnts.Add(roundCirLeft.ObjectId);

                                acBlkTblRec.AppendEntity(roundCirRight);
                                Trans.AddNewlyCreatedDBObject(roundCirRight, true);
                                allPoperEnts.Add(roundCirRight.ObjectId);
                                break;
                            }
                        }

                    }
                    //далее строим код 140 и 141 - все то же самое, что и 130 и 131, только кружочки сверху
                    if (myCodeTxt.TextString == "140")
                    {
                        using (Line AllVertLine = new Line(myCodeTxt.Position, new Point3d(myCodeTxt.Position.X, vspomLine4.StartPoint.Y, 0)))
                        {
                            //сначала находим вершину верха коммуникации
                            Point3dCollection intCollection = new Point3dCollection();
                            AllVertLine.IntersectWith(poperLine, intersectType: Intersect.OnBothOperands, intCollection, IntPtr.Zero, IntPtr.Zero);
                            Point3d point3D_INTER = intCollection[0]; //проблема! (при чтении традиционных рор-файлов
                            Point3d point3D_Down = new Point3d(myCodeTxt.Position.X, gorizontLine.StartPoint.Y, 0);
                            Point3d point3D_Up = new Point3d(point3D_INTER.X, point3D_INTER.Y + vertScaleKoef * deltaH, 0);
                            Line lineLep = new Line(point3D_Down, point3D_Up);//строится вертикальная линия от горизонта до верхней точки с превышением
                            acBlkTblRec.AppendEntity(lineLep);
                            Trans.AddNewlyCreatedDBObject(lineLep, true);
                            allPoperEnts.Add(lineLep.ObjectId);

                            Point3d point3DLeftLep = new Point3d(point3D_Up.X - 5, point3D_Up.Y, 0);
                            Point3d point3DRightLep = new Point3d(point3D_Up.X + 5, point3D_Up.Y, 0);
                            Line lineLepGor = new Line(point3DLeftLep, point3DRightLep);//горизонтальная линия значка ЛЭП
                            acBlkTblRec.AppendEntity(lineLepGor);
                            Trans.AddNewlyCreatedDBObject(lineLepGor, true);
                            allPoperEnts.Add(lineLepGor.ObjectId);

                            Point3d point3DleftCir = new Point3d(point3D_Up.X - 4, point3D_Up.Y + 1, 0);
                            Point3d point3DRightCir = new Point3d(point3D_Up.X + 4, point3D_Up.Y + 1, 0);
                            Circle roundCirLeft = new Circle(point3DleftCir, Vector3d.ZAxis, 0.5);
                            Circle roundCirRight = new Circle(point3DRightCir, Vector3d.ZAxis, 0.5);//слева и справа рисуем окружности значка ЛЭП
                            acBlkTblRec.AppendEntity(roundCirLeft);
                            Trans.AddNewlyCreatedDBObject(roundCirLeft, true);
                            allPoperEnts.Add(roundCirLeft.ObjectId);

                            acBlkTblRec.AppendEntity(roundCirRight);
                            Trans.AddNewlyCreatedDBObject(roundCirRight, true);
                            allPoperEnts.Add(roundCirRight.ObjectId);
                        }

                    }
                    if (myCodeTxt.TextString == "141")//для случая, когда задана абсолютная отметка верха коммуникации,
                                                      //строим условный знак до точки из коллекции colIdPodzemkaAbsolut
                    {
                        //ищем в коллекции colIdPodzemkaAbsolut соответствующую точку
                        for (int i = 0; i < colIdPodzemkaAbsolut.Count; i++)
                        {
                            DBPoint myPodzemPt = (DBPoint)Trans.GetObject(colIdPodzemkaAbsolut[i], OpenMode.ForWrite);
                            if (myPodzemPt.Position.X == myCodeTxt.Position.X)
                            {
                                Point3d point3D_Up = new Point3d(myPodzemPt.Position.X, myPodzemPt.Position.Y, 0);//найдена верхняя абсолютная точка
                                Point3d point3D_Down = new Point3d(myCodeTxt.Position.X, gorizontLine.StartPoint.Y, 0);
                                Line lineLep = new Line(point3D_Down, point3D_Up);//строится вертикальная линия от горизонта до верхней точки с превышением
                                acBlkTblRec.AppendEntity(lineLep);
                                Trans.AddNewlyCreatedDBObject(lineLep, true);
                                allPoperEnts.Add(lineLep.ObjectId);

                                Point3d point3DLeftLep = new Point3d(point3D_Up.X - 5, point3D_Up.Y, 0);
                                Point3d point3DRightLep = new Point3d(point3D_Up.X + 5, point3D_Up.Y, 0);
                                Line lineLepGor = new Line(point3DLeftLep, point3DRightLep);//горизонтальная линия значка ЛЭП
                                acBlkTblRec.AppendEntity(lineLepGor);
                                Trans.AddNewlyCreatedDBObject(lineLepGor, true);
                                allPoperEnts.Add(lineLepGor.ObjectId);

                                Point3d point3DleftCir = new Point3d(point3D_Up.X - 4, point3D_Up.Y + 1, 0);
                                Point3d point3DRightCir = new Point3d(point3D_Up.X + 4, point3D_Up.Y + 1, 0);
                                Circle roundCirLeft = new Circle(point3DleftCir, Vector3d.ZAxis, 0.5);
                                Circle roundCirRight = new Circle(point3DRightCir, Vector3d.ZAxis, 0.5);//слева и справа рисуем окружности значка ЛЭП
                                acBlkTblRec.AppendEntity(roundCirLeft);
                                Trans.AddNewlyCreatedDBObject(roundCirLeft, true);
                                allPoperEnts.Add(roundCirLeft.ObjectId);

                                acBlkTblRec.AppendEntity(roundCirRight);
                                Trans.AddNewlyCreatedDBObject(roundCirRight, true);
                                allPoperEnts.Add(roundCirRight.ObjectId);
                                break;
                            }
                        }

                    }
                }

                //строим условный знак для кодов 2 и 3 (стена здания)
                if (myCodeTxt.TextString == "2" || myCodeTxt.TextString == "3")
                {
                    using (Line AllVertLine = new Line(myCodeTxt.Position, new Point3d(myCodeTxt.Position.X, vspomLine4.StartPoint.Y, 0)))
                    {
                        //находим точку пересечения линии кода с линией профиля - оно все равно получится в существующей вершине
                        Point3dCollection intCollection = new Point3dCollection();
                        AllVertLine.IntersectWith(poperLine, intersectType: Intersect.OnBothOperands, intCollection, IntPtr.Zero, IntPtr.Zero);
                        Point3d point3D_INTER = intCollection[0]; //проблема! (при чтении традиционных рор-файлов

                        Point3d point3D_Up = new Point3d(point3D_INTER.X, point3D_INTER.Y + 24, 0);
                        Line lineLep = new Line(point3D_INTER, point3D_Up);//строится вертикальная линия от профиля до верхней точки (4.8х верт.масштаб)
                        acBlkTblRec.AppendEntity(lineLep);
                        Trans.AddNewlyCreatedDBObject(lineLep, true);
                        allPoperEnts.Add(lineLep.ObjectId);
                        if (myCodeTxt.TextString == "2") // при коде 2 косая палочка: лево-низ - право-верх
                        {
                            Point3d point3DLeftLep = new Point3d(point3D_Up.X - 1, point3D_Up.Y - 1, 0);
                            Point3d point3DRightLep = new Point3d(point3D_Up.X + 1, point3D_Up.Y + 1, 0);
                            Line lineLepGor = new Line(point3DLeftLep, point3DRightLep);//горизонтальная линия значка ЛЭП
                            acBlkTblRec.AppendEntity(lineLepGor);
                            Trans.AddNewlyCreatedDBObject(lineLepGor, true);
                            allPoperEnts.Add(lineLepGor.ObjectId);
                        }
                        else if (myCodeTxt.TextString == "3") // при коде 3 косая палочка: лево-верх - право-низ
                        {
                            Point3d point3DLeftLep = new Point3d(point3D_Up.X - 1, point3D_Up.Y + 1, 0);
                            Point3d point3DRightLep = new Point3d(point3D_Up.X + 1, point3D_Up.Y - 1, 0);
                            Line lineLepGor = new Line(point3DLeftLep, point3DRightLep);//горизонтальная линия значка ЛЭП
                            acBlkTblRec.AppendEntity(lineLepGor);
                            Trans.AddNewlyCreatedDBObject(lineLepGor, true);
                            allPoperEnts.Add(lineLepGor.ObjectId);
                        }


                    }

                }
                //строим условный знак для кодов 4 - 9 (заборы)
                if (myCodeTxt.TextString == "4" || myCodeTxt.TextString == "5" || myCodeTxt.TextString == "6" || myCodeTxt.TextString == "7" || myCodeTxt.TextString == "8" || myCodeTxt.TextString == "9")
                {
                    using (Line AllVertLine = new Line(myCodeTxt.Position, new Point3d(myCodeTxt.Position.X, vspomLine4.StartPoint.Y, 0)))
                    {
                        //находим точку пересечения линии кода с линией профиля - оно все равно получится в существующей вершине
                        Point3dCollection intCollection = new Point3dCollection();
                        AllVertLine.IntersectWith(poperLine, intersectType: Intersect.OnBothOperands, intCollection, IntPtr.Zero, IntPtr.Zero);
                        Point3d point3D_INTER = intCollection[0]; //проблема! (при чтении традиционных рор-файлов

                        if (myCodeTxt.TextString == "4") // при коде 4 рисуем от точки пересечения УЗ забора бетонного больше метра
                        {
                            Point3d point3D_1 = new Point3d(point3D_INTER.X - 0.4, point3D_INTER.Y, 0);
                            Point3d point3D_2 = new Point3d(point3D_INTER.X + 0.4, point3D_INTER.Y, 0);
                            Line line1 = new Line(point3D_1, point3D_2);//горизонтальная нижняя линия значка забора
                            acBlkTblRec.AppendEntity(line1);
                            Trans.AddNewlyCreatedDBObject(line1, true);
                            allPoperEnts.Add(line1.ObjectId);
                            Point3d point3D_3 = new Point3d(point3D_INTER.X + 0.4, point3D_INTER.Y + 12, 0);
                            Line line2 = new Line(point3D_2, point3D_3);//вертикальная правая длинная линия значка забора
                            acBlkTblRec.AppendEntity(line2);
                            Trans.AddNewlyCreatedDBObject(line2, true);
                            allPoperEnts.Add(line2.ObjectId);
                            Point3d point3D_4 = new Point3d(point3D_INTER.X - 0.4, point3D_INTER.Y + 12, 0);
                            Line line3 = new Line(point3D_3, point3D_4);//горизонтальная верхняя линия значка забора
                            acBlkTblRec.AppendEntity(line3);
                            Trans.AddNewlyCreatedDBObject(line3, true);
                            allPoperEnts.Add(line3.ObjectId);
                            Line line4 = new Line(point3D_1, point3D_4);//вертикальная левая длинная линия значка забора
                            acBlkTblRec.AppendEntity(line4);
                            Trans.AddNewlyCreatedDBObject(line4, true);
                            allPoperEnts.Add(line4.ObjectId);
                            Point3d point3D_5 = new Point3d(point3D_INTER.X - 0.4, point3D_INTER.Y + 4, 0);
                            Point3d point3D_6 = new Point3d(point3D_INTER.X + 0.4, point3D_INTER.Y + 4, 0);
                            Line line5 = new Line(point3D_5, point3D_6);//горизонтальная средняя линия значка забора
                            acBlkTblRec.AppendEntity(line5);
                            Trans.AddNewlyCreatedDBObject(line5, true);
                            allPoperEnts.Add(line5.ObjectId);
                            Point3d point3D_7 = new Point3d(point3D_INTER.X + 0.4, point3D_INTER.Y + 7.5, 0);
                            Point3d point3D_8 = new Point3d(point3D_INTER.X + 1.2, point3D_INTER.Y + 7.5, 0);
                            Line line6 = new Line(point3D_7, point3D_8);
                            acBlkTblRec.AppendEntity(line6);
                            Trans.AddNewlyCreatedDBObject(line6, true);
                            allPoperEnts.Add(line6.ObjectId);
                            Point3d point3D_9 = new Point3d(point3D_INTER.X + 1.2, point3D_INTER.Y + 8.5, 0);
                            Line line7 = new Line(point3D_8, point3D_9);
                            acBlkTblRec.AppendEntity(line7);
                            Trans.AddNewlyCreatedDBObject(line7, true);
                            allPoperEnts.Add(line7.ObjectId);
                            Point3d point3D_10 = new Point3d(point3D_INTER.X + 0.4, point3D_INTER.Y + 8.5, 0);
                            Line line8 = new Line(point3D_9, point3D_10);
                            acBlkTblRec.AppendEntity(line8);
                            Trans.AddNewlyCreatedDBObject(line8, true);
                            allPoperEnts.Add(line8.ObjectId);

                        }
                        else if (myCodeTxt.TextString == "5") // бетонный низкий забор
                        {
                            Point3d point3D_1 = new Point3d(point3D_INTER.X - 0.4, point3D_INTER.Y, 0);
                            Point3d point3D_2 = new Point3d(point3D_INTER.X + 0.4, point3D_INTER.Y, 0);
                            Line line1 = new Line(point3D_1, point3D_2);//горизонтальная нижняя линия значка забора
                            acBlkTblRec.AppendEntity(line1);
                            Trans.AddNewlyCreatedDBObject(line1, true);
                            allPoperEnts.Add(line1.ObjectId);
                            Point3d point3D_3 = new Point3d(point3D_INTER.X + 0.4, point3D_INTER.Y + 12, 0);
                            Line line2 = new Line(point3D_2, point3D_3);//вертикальная правая длинная линия значка забора
                            acBlkTblRec.AppendEntity(line2);
                            Trans.AddNewlyCreatedDBObject(line2, true);
                            allPoperEnts.Add(line2.ObjectId);
                            Point3d point3D_4 = new Point3d(point3D_INTER.X - 0.4, point3D_INTER.Y + 12, 0);
                            Line line3 = new Line(point3D_3, point3D_4);//горизонтальная верхняя линия значка забора
                            acBlkTblRec.AppendEntity(line3);
                            Trans.AddNewlyCreatedDBObject(line3, true);
                            allPoperEnts.Add(line3.ObjectId);
                            Line line4 = new Line(point3D_1, point3D_4);//вертикальная левая длинная линия значка забора
                            acBlkTblRec.AppendEntity(line4);
                            Trans.AddNewlyCreatedDBObject(line4, true);
                            allPoperEnts.Add(line4.ObjectId);
                            Point3d point3D_5 = new Point3d(point3D_INTER.X - 0.4, point3D_INTER.Y + 4, 0);
                            Point3d point3D_6 = new Point3d(point3D_INTER.X + 0.4, point3D_INTER.Y + 4, 0);
                            Line line5 = new Line(point3D_5, point3D_6);//горизонтальная средняя линия значка забора
                            acBlkTblRec.AppendEntity(line5);
                            Trans.AddNewlyCreatedDBObject(line5, true);
                            allPoperEnts.Add(line5.ObjectId);
                            Point3d point3D_7 = new Point3d(point3D_INTER.X - 0.4, point3D_INTER.Y + 3.4, 0);
                            Point3d point3D_8 = new Point3d(point3D_INTER.X + 0.4, point3D_INTER.Y + 3.4, 0);
                            Line line6 = new Line(point3D_7, point3D_8);//
                            acBlkTblRec.AppendEntity(line6);
                            Trans.AddNewlyCreatedDBObject(line6, true);
                            allPoperEnts.Add(line6.ObjectId);
                            Point3d point3D_9 = new Point3d(point3D_INTER.X - 0.4, point3D_INTER.Y + 11.4, 0);
                            Point3d point3D_10 = new Point3d(point3D_INTER.X + 0.4, point3D_INTER.Y + 11.4, 0);
                            Line line7 = new Line(point3D_9, point3D_10);//
                            acBlkTblRec.AppendEntity(line7);
                            Trans.AddNewlyCreatedDBObject(line7, true);
                            allPoperEnts.Add(line7.ObjectId);
                        }
                        else if (myCodeTxt.TextString == "6") // металлический высокий забор
                        {
                            Point3d point3D_1 = new Point3d(point3D_INTER.X, point3D_INTER.Y + 4.6, 0);
                            Line line1 = new Line(point3D_INTER, point3D_1);//первая вертикальная линия значка забора
                            acBlkTblRec.AppendEntity(line1);
                            Trans.AddNewlyCreatedDBObject(line1, true);
                            allPoperEnts.Add(line1.ObjectId);
                            Point3d point3D_2 = new Point3d(point3D_INTER.X, point3D_INTER.Y + 5.4, 0);
                            Point3d point3D_3 = new Point3d(point3D_INTER.X, point3D_INTER.Y + 9.6, 0);
                            Line line2 = new Line(point3D_2, point3D_3);//вторая вертикальная линия значка забора
                            acBlkTblRec.AppendEntity(line2);
                            Trans.AddNewlyCreatedDBObject(line2, true);
                            allPoperEnts.Add(line2.ObjectId);
                            Point3d point3D_4 = new Point3d(point3D_INTER.X + 0.4, point3D_INTER.Y + 5, 0);
                            Point3d point3D_5 = new Point3d(point3D_INTER.X + 0.9, point3D_INTER.Y + 5, 0);
                            Line line3 = new Line(point3D_4, point3D_5);//горизонтальная верхняя линия значка забора
                            acBlkTblRec.AppendEntity(line3);
                            Trans.AddNewlyCreatedDBObject(line3, true);
                            allPoperEnts.Add(line3.ObjectId);
                            Point3d point3D_6 = new Point3d(point3D_INTER.X + 0.4, point3D_INTER.Y + 10, 0);
                            Point3d point3D_7 = new Point3d(point3D_INTER.X + 0.9, point3D_INTER.Y + 10, 0);
                            Line line4 = new Line(point3D_6, point3D_7);
                            acBlkTblRec.AppendEntity(line4);
                            Trans.AddNewlyCreatedDBObject(line4, true);
                            allPoperEnts.Add(line4.ObjectId);
                            //рисуем окружности в заборе
                            Point3d point3D_8 = new Point3d(point3D_INTER.X, point3D_INTER.Y + 5, 0);
                            Circle circle1 = new Circle(point3D_8, Vector3d.ZAxis, 0.1);
                            acBlkTblRec.AppendEntity(circle1);
                            Trans.AddNewlyCreatedDBObject(circle1, true);
                            allPoperEnts.Add(circle1.ObjectId);
                            Circle circle2 = new Circle(point3D_8, Vector3d.ZAxis, 0.2);
                            acBlkTblRec.AppendEntity(circle2);
                            Trans.AddNewlyCreatedDBObject(circle2, true);
                            allPoperEnts.Add(circle2.ObjectId);
                            Circle circle3 = new Circle(point3D_8, Vector3d.ZAxis, 0.3);
                            acBlkTblRec.AppendEntity(circle3);
                            Trans.AddNewlyCreatedDBObject(circle3, true);
                            allPoperEnts.Add(circle3.ObjectId);
                            Circle circle4 = new Circle(point3D_8, Vector3d.ZAxis, 0.4);
                            acBlkTblRec.AppendEntity(circle4);
                            Trans.AddNewlyCreatedDBObject(circle4, true);
                            allPoperEnts.Add(circle4.ObjectId);

                            Point3d point3D_9 = new Point3d(point3D_INTER.X, point3D_INTER.Y + 10, 0);
                            Circle circle5 = new Circle(point3D_9, Vector3d.ZAxis, 0.1);
                            acBlkTblRec.AppendEntity(circle5);
                            Trans.AddNewlyCreatedDBObject(circle5, true);
                            allPoperEnts.Add(circle5.ObjectId);
                            Circle circle6 = new Circle(point3D_9, Vector3d.ZAxis, 0.2);
                            acBlkTblRec.AppendEntity(circle6);
                            Trans.AddNewlyCreatedDBObject(circle6, true);
                            allPoperEnts.Add(circle6.ObjectId);
                            Circle circle7 = new Circle(point3D_9, Vector3d.ZAxis, 0.3);
                            acBlkTblRec.AppendEntity(circle7);
                            Trans.AddNewlyCreatedDBObject(circle7, true);
                            allPoperEnts.Add(circle7.ObjectId);
                            Circle circle8 = new Circle(point3D_9, Vector3d.ZAxis, 0.4);
                            acBlkTblRec.AppendEntity(circle8);
                            Trans.AddNewlyCreatedDBObject(circle8, true);
                            allPoperEnts.Add(circle8.ObjectId);
                        }
                        else if (myCodeTxt.TextString == "7") // металлический низкий забор - без боковых палочек
                        {
                            Point3d point3D_1 = new Point3d(point3D_INTER.X, point3D_INTER.Y + 4.6, 0);
                            Line line1 = new Line(point3D_INTER, point3D_1);//первая вертикальная линия значка забора
                            acBlkTblRec.AppendEntity(line1);
                            Trans.AddNewlyCreatedDBObject(line1, true);
                            allPoperEnts.Add(line1.ObjectId);
                            Point3d point3D_2 = new Point3d(point3D_INTER.X, point3D_INTER.Y + 5.4, 0);
                            Point3d point3D_3 = new Point3d(point3D_INTER.X, point3D_INTER.Y + 9.6, 0);
                            Line line2 = new Line(point3D_2, point3D_3);//вторая вертикальная линия значка забора
                            acBlkTblRec.AppendEntity(line2);
                            Trans.AddNewlyCreatedDBObject(line2, true);
                            allPoperEnts.Add(line2.ObjectId);

                            //рисуем окружности в заборе
                            Point3d point3D_8 = new Point3d(point3D_INTER.X, point3D_INTER.Y + 5, 0);
                            Circle circle1 = new Circle(point3D_8, Vector3d.ZAxis, 0.1);
                            acBlkTblRec.AppendEntity(circle1);
                            Trans.AddNewlyCreatedDBObject(circle1, true);
                            allPoperEnts.Add(circle1.ObjectId);
                            Circle circle2 = new Circle(point3D_8, Vector3d.ZAxis, 0.2);
                            acBlkTblRec.AppendEntity(circle2);
                            Trans.AddNewlyCreatedDBObject(circle2, true);
                            allPoperEnts.Add(circle2.ObjectId);
                            Circle circle3 = new Circle(point3D_8, Vector3d.ZAxis, 0.3);
                            acBlkTblRec.AppendEntity(circle3);
                            Trans.AddNewlyCreatedDBObject(circle3, true);
                            allPoperEnts.Add(circle3.ObjectId);
                            Circle circle4 = new Circle(point3D_8, Vector3d.ZAxis, 0.4);
                            acBlkTblRec.AppendEntity(circle4);
                            Trans.AddNewlyCreatedDBObject(circle4, true);
                            allPoperEnts.Add(circle4.ObjectId);

                            Point3d point3D_9 = new Point3d(point3D_INTER.X, point3D_INTER.Y + 10, 0);
                            Circle circle5 = new Circle(point3D_9, Vector3d.ZAxis, 0.1);
                            acBlkTblRec.AppendEntity(circle5);
                            Trans.AddNewlyCreatedDBObject(circle5, true);
                            allPoperEnts.Add(circle5.ObjectId);
                            Circle circle6 = new Circle(point3D_9, Vector3d.ZAxis, 0.2);
                            acBlkTblRec.AppendEntity(circle6);
                            Trans.AddNewlyCreatedDBObject(circle6, true);
                            allPoperEnts.Add(circle6.ObjectId);
                            Circle circle7 = new Circle(point3D_9, Vector3d.ZAxis, 0.3);
                            acBlkTblRec.AppendEntity(circle7);
                            Trans.AddNewlyCreatedDBObject(circle7, true);
                            allPoperEnts.Add(circle7.ObjectId);
                            Circle circle8 = new Circle(point3D_9, Vector3d.ZAxis, 0.4);
                            acBlkTblRec.AppendEntity(circle8);
                            Trans.AddNewlyCreatedDBObject(circle8, true);
                            allPoperEnts.Add(circle8.ObjectId);
                        }
                        else if (myCodeTxt.TextString == "8") // деревянный забор
                        {

                            Point3d point3D_2 = new Point3d(point3D_INTER.X, point3D_INTER.Y + 10, 0);
                            Line line1 = new Line(point3D_INTER, point3D_2);//вертикальная линия значка забора
                            acBlkTblRec.AppendEntity(line1);
                            Trans.AddNewlyCreatedDBObject(line1, true);
                            allPoperEnts.Add(line1.ObjectId);

                            Point3d point3D_3 = new Point3d(point3D_INTER.X, point3D_INTER.Y + 2.6, 0);//_________________________со всех значков подземки выше убрать зависимость от масштаба!!!
                            Point3d point3D_4 = new Point3d(point3D_INTER.X + 0.8, point3D_INTER.Y + 2.6, 0);
                            Line line2 = new Line(point3D_3, point3D_4);
                            acBlkTblRec.AppendEntity(line2);
                            Trans.AddNewlyCreatedDBObject(line2, true);
                            allPoperEnts.Add(line2.ObjectId);
                            Point3d point3D_5 = new Point3d(point3D_INTER.X, point3D_INTER.Y + 7.6, 0);
                            Point3d point3D_6 = new Point3d(point3D_INTER.X + 0.8, point3D_INTER.Y + 7.6, 0);
                            Line line3 = new Line(point3D_5, point3D_6);
                            acBlkTblRec.AppendEntity(line3);
                            Trans.AddNewlyCreatedDBObject(line3, true);
                            allPoperEnts.Add(line3.ObjectId);
                        }
                        else if (myCodeTxt.TextString == "9") // рабица забор
                        {

                            Point3d point3D_2 = new Point3d(point3D_INTER.X, point3D_INTER.Y + 10, 0);
                            Line line1 = new Line(point3D_INTER, point3D_2);//вертикальная линия значка забора
                            acBlkTblRec.AppendEntity(line1);
                            Trans.AddNewlyCreatedDBObject(line1, true);
                            allPoperEnts.Add(line1.ObjectId);

                            Point3d point3D_3 = new Point3d(point3D_INTER.X, point3D_INTER.Y + 3, 0);
                            Point3d point3D_4 = new Point3d(point3D_INTER.X + 0.55, point3D_INTER.Y + 2.45, 0);
                            Line line2 = new Line(point3D_3, point3D_4);
                            acBlkTblRec.AppendEntity(line2);
                            Trans.AddNewlyCreatedDBObject(line2, true);
                            allPoperEnts.Add(line2.ObjectId);
                            Point3d point3D_5 = new Point3d(point3D_INTER.X + 0.55, point3D_INTER.Y + 3.55, 0);
                            Line line3 = new Line(point3D_3, point3D_5);
                            acBlkTblRec.AppendEntity(line3);
                            Trans.AddNewlyCreatedDBObject(line3, true);
                            allPoperEnts.Add(line3.ObjectId);

                            Point3d point3D_6 = new Point3d(point3D_INTER.X, point3D_INTER.Y + 8, 0);
                            Point3d point3D_7 = new Point3d(point3D_INTER.X - 0.55, point3D_INTER.Y + 7.45, 0);
                            Line line4 = new Line(point3D_6, point3D_7);
                            acBlkTblRec.AppendEntity(line4);
                            Trans.AddNewlyCreatedDBObject(line4, true);
                            allPoperEnts.Add(line4.ObjectId);
                            Point3d point3D_8 = new Point3d(point3D_INTER.X - 0.55, point3D_INTER.Y + 8.55, 0);
                            Line line5 = new Line(point3D_6, point3D_8);
                            acBlkTblRec.AppendEntity(line5);
                            Trans.AddNewlyCreatedDBObject(line5, true);
                            allPoperEnts.Add(line5.ObjectId);

                        }

                    }

                }
                //делаем код 11 ряд кустов
                if (myCodeTxt.TextString == "11")
                {
                    using (Line AllVertLine = new Line(myCodeTxt.Position, new Point3d(myCodeTxt.Position.X, vspomLine4.StartPoint.Y, 0)))
                    {
                        //находим точку пересечения линии кода с линией профиля - оно все равно получится в существующей вершине
                        Point3dCollection intCollection = new Point3dCollection();
                        AllVertLine.IntersectWith(poperLine, intersectType: Intersect.OnBothOperands, intCollection, IntPtr.Zero, IntPtr.Zero);
                        Point3d point3D_INTER = intCollection[0];
                        Point3d point1 = new Point3d(point3D_INTER.X, point3D_INTER.Y + 0.4, 0);
                        Circle circle1 = new Circle(point1, Vector3d.ZAxis, 0.4);
                        acBlkTblRec.AppendEntity(circle1);
                        Trans.AddNewlyCreatedDBObject(circle1, true);
                        allPoperEnts.Add(circle1.ObjectId);

                        Point3d point2 = new Point3d(point3D_INTER.X, point3D_INTER.Y + 3, 0);
                        Circle circle2 = new Circle(point2, Vector3d.ZAxis, 0.07);
                        acBlkTblRec.AppendEntity(circle2);
                        Trans.AddNewlyCreatedDBObject(circle2, true);
                        allPoperEnts.Add(circle2.ObjectId);
                        Circle circle3 = new Circle(point2, Vector3d.ZAxis, 0.14);
                        acBlkTblRec.AppendEntity(circle3);
                        Trans.AddNewlyCreatedDBObject(circle3, true);
                        allPoperEnts.Add(circle3.ObjectId);
                        Circle circle4 = new Circle(point2, Vector3d.ZAxis, 0.21);
                        acBlkTblRec.AppendEntity(circle4);
                        Trans.AddNewlyCreatedDBObject(circle4, true);
                        allPoperEnts.Add(circle4.ObjectId);

                        Point3d point3 = new Point3d(point3D_INTER.X, point3D_INTER.Y + 5.6, 0);
                        Circle circle5 = new Circle(point3, Vector3d.ZAxis, 0.4);
                        acBlkTblRec.AppendEntity(circle5);
                        Trans.AddNewlyCreatedDBObject(circle5, true);
                        allPoperEnts.Add(circle5.ObjectId);

                        Point3d point4 = new Point3d(point3D_INTER.X, point3D_INTER.Y + 8.2, 0);
                        Circle circle6 = new Circle(point4, Vector3d.ZAxis, 0.07);
                        acBlkTblRec.AppendEntity(circle6);
                        Trans.AddNewlyCreatedDBObject(circle6, true);
                        allPoperEnts.Add(circle6.ObjectId);
                        Circle circle7 = new Circle(point4, Vector3d.ZAxis, 0.14);
                        acBlkTblRec.AppendEntity(circle7);
                        Trans.AddNewlyCreatedDBObject(circle7, true);
                        allPoperEnts.Add(circle7.ObjectId);
                        Circle circle8 = new Circle(point4, Vector3d.ZAxis, 0.21);
                        acBlkTblRec.AppendEntity(circle8);
                        Trans.AddNewlyCreatedDBObject(circle8, true);
                        allPoperEnts.Add(circle8.ObjectId);
                    }
                }
                //делаем код 15 ряд деревьев
                if (myCodeTxt.TextString == "15")
                {
                    using (Line AllVertLine = new Line(myCodeTxt.Position, new Point3d(myCodeTxt.Position.X, vspomLine4.StartPoint.Y, 0)))
                    {
                        //находим точку пересечения линии кода с линией профиля - оно все равно получится в существующей вершине
                        Point3dCollection intCollection = new Point3dCollection();
                        AllVertLine.IntersectWith(poperLine, intersectType: Intersect.OnBothOperands, intCollection, IntPtr.Zero, IntPtr.Zero);
                        Point3d point3D_INTER = intCollection[0];
                        Polyline treePolyline = new Polyline();// создаем пустую полилинию, в которую попробуем добавлять элементы дерева
                        Point2d point1 = new Point2d(point3D_INTER.X + 0.5, point3D_INTER.Y);
                        treePolyline.AddVertexAt(0, point1, 0, 0, 0);
                        treePolyline.AddVertexAt(1, new Point2d(point3D_INTER.X, point3D_INTER.Y), 0, 0, 0);
                        acBlkTblRec.AppendEntity(treePolyline);
                        Trans.AddNewlyCreatedDBObject(treePolyline, true);
                        allPoperEnts.Add(treePolyline.ObjectId);
                        Point2d point2 = new Point2d(point3D_INTER.X, point3D_INTER.Y + 1);
                        treePolyline.AddVertexAt(2, point2, 0, 0, 0);
                        Point2d point3 = new Point2d(point3D_INTER.X + 0.5131, point3D_INTER.Y + 1);
                        treePolyline.AddVertexAt(3, point3, 0.5, 0, 0);
                        Point2d point4 = new Point2d(point3D_INTER.X + 0.73, point3D_INTER.Y + 1.95);
                        treePolyline.AddVertexAt(4, point4, 0.5, 0, 0);
                        Point2d point5 = new Point2d(point3D_INTER.X + 0.46, point3D_INTER.Y + 2.81);
                        treePolyline.AddVertexAt(5, point5, 0.5, 0, 0);
                        Point2d point6 = new Point2d(point3D_INTER.X, point3D_INTER.Y + 3.5);
                        treePolyline.AddVertexAt(6, point6, 0.5, 0, 0);
                        Point2d point7 = new Point2d(point3D_INTER.X - 0.46, point3D_INTER.Y + 2.81);
                        treePolyline.AddVertexAt(7, point7, 0.5, 0, 0);
                        Point2d point8 = new Point2d(point3D_INTER.X - 0.73, point3D_INTER.Y + 1.95);
                        treePolyline.AddVertexAt(8, point8, 0.5, 0, 0);
                        Point2d point9 = new Point2d(point3D_INTER.X - 0.5131, point3D_INTER.Y + 1);
                        treePolyline.AddVertexAt(9, point9, 0, 0, 0);
                        treePolyline.AddVertexAt(10, point2, 0, 0, 0);
                    }
                }
                if (myCodeTxt.TextString == "501")//создаем отдельную выноску для кода ИССО (по-разному для стороны лево и право,
                                                  //строим условный знак до точки из коллекции colIdPodzemkaAbsolut
                {
                    //ищем в коллекции colIdPodzemkaAbsolut соответствующую точку
                    for (int i = 0; i < colIdPodzemkaAbsolut.Count; i++)
                    {
                        DBPoint myPodzemPt = (DBPoint)Trans.GetObject(colIdPodzemkaAbsolut[i], OpenMode.ForWrite);

                        if (myPodzemPt.Position.X == myCodeTxt.Position.X)
                        {
                            Point2d point2D_Up = new Point2d(myPodzemPt.Position.X, myPodzemPt.Position.Y);//найдена верхняя абсолютная точка
                                                                                                           //строим от нее выноску-стрелочку
                            double rasstIsso = 0;
                            Polyline polylineVynos = new Polyline();
                            acBlkTblRec.AppendEntity(polylineVynos);
                            Trans.AddNewlyCreatedDBObject(polylineVynos, true);
                            allPoperEnts.Add(polylineVynos.ObjectId);
                            polylineVynos.AddVertexAt(0, point2D_Up, 0, 0, 0.5);
                            DBText dBTextOpisanieIsso = new DBText();
                            DBText dBTextOtmetIsso = new DBText();

                            for (int j = 0; j < colIdOpisanie.Count; ++j)//находим в коллекции описание точки ИССО и вставляем его пока в начальную точку
                            {
                                dBTextOpisanieIsso = (DBText)Trans.GetObject(colIdOpisanie[j], OpenMode.ForWrite);
                                if (dBTextOpisanieIsso.Position.X == myPodzemPt.Position.X)
                                {
                                    dBTextOpisanieIsso.Position = myPodzemPt.Position;
                                    dBTextOpisanieIsso.Rotation = 0;
                                    dBTextOpisanieIsso.ColorIndex = 256;
                                    break;
                                }
                            }

                            for (int j = 0; j < colIdPodzemkaRasst.Count; ++j)//находим в коллекции отметку точки ИССО и вставляем ее пока в начальную точку
                            {
                                dBTextOtmetIsso = (DBText)Trans.GetObject(colIdPodzemkaRasst[j], OpenMode.ForWrite);
                                if (dBTextOtmetIsso.Position.X == myPodzemPt.Position.X)
                                {
                                    dBTextOtmetIsso.Position = myPodzemPt.Position;
                                    dBTextOtmetIsso.Rotation = 0;
                                    dBTextOtmetIsso.ColorIndex = 256;
                                    allPoperEnts.Add(colIdPodzemkaRasst[j]);
                                    colIdPodzemkaRasst.Remove(colIdPodzemkaRasst[j]);
                                    break;
                                }

                            }

                            for (int j = 0; j < colIdRasstIsso.Count; ++j) //определяем, слева или справа от 0-точки находится данный код
                            {
                                DBText dBTextRasstIsso = (DBText)Trans.GetObject(colIdRasstIsso[j], OpenMode.ForWrite);
                                if (dBTextRasstIsso.Position.X == myPodzemPt.Position.X)
                                {
                                    rasstIsso = Convert.ToDouble(dBTextRasstIsso.TextString);
                                    // ed.WriteMessage($"Получено значение rasstIsso= {rasstIsso}\n");
                                    break;
                                }

                            }
                            if (rasstIsso > 0) //рисуем стрелку-выноску вправо и переставляем тексты с отметками и описаниями
                            {
                                Point2d point2D_1 = new Point2d(point2D_Up.X + 0.817, point2D_Up.Y + 0.946);
                                polylineVynos.AddVertexAt(1, point2D_1, 0, 0, 0);
                                Point2d point2D_2 = new Point2d(point2D_Up.X + 4.085, point2D_Up.Y + 4.7302);
                                polylineVynos.AddVertexAt(2, point2D_2, 0, 0, 0);
                                Point2d point2D_3 = new Point2d(point2D_Up.X + 19.085, point2D_Up.Y + 4.7302);
                                polylineVynos.AddVertexAt(3, point2D_3, 0, 0, 0);
                                Point3d point3DOpis = new Point3d(point2D_Up.X + 4.585, point2D_Up.Y + 5.0302, 0);
                                dBTextOpisanieIsso.Position = point3DOpis;
                                Point3d point3DOtm = new Point3d(point2D_Up.X + 4.585, point2D_2.Y - 0.3 - dBTextOtmetIsso.Height, 0);
                                dBTextOtmetIsso.Position = point3DOtm;
                            }
                            else
                            {
                                Point2d point2D_1 = new Point2d(point2D_Up.X - 0.817, point2D_Up.Y + 0.946);
                                polylineVynos.AddVertexAt(1, point2D_1, 0, 0, 0);
                                Point2d point2D_2 = new Point2d(point2D_Up.X - 4.085, point2D_Up.Y + 4.7302);
                                polylineVynos.AddVertexAt(2, point2D_2, 0, 0, 0);
                                Point2d point2D_3 = new Point2d(point2D_Up.X - 19.085, point2D_Up.Y + 4.7302);
                                polylineVynos.AddVertexAt(3, point2D_3, 0, 0, 0);
                                Point3d point3DOpis = new Point3d(point2D_2.X - (dBTextOpisanieIsso.GeometricExtents.MaxPoint.X - dBTextOpisanieIsso.GeometricExtents.MinPoint.X), point2D_Up.Y + 5.0302, 0);
                                dBTextOpisanieIsso.Position = point3DOpis;
                                Point3d point3DOtm = new Point3d(point2D_2.X - (dBTextOtmetIsso.GeometricExtents.MaxPoint.X - dBTextOtmetIsso.GeometricExtents.MinPoint.X), point2D_2.Y - 0.3 - dBTextOtmetIsso.Height, 0);
                                dBTextOtmetIsso.Position = point3DOtm;
                            }
                            break;
                        }
                    }

                }

            }
            //---------для каждого кода пути "1" рисуем сверху междопутье-----------------------
            if (pointsFinalPunktir.Count > 1)
            {

                CommonMethods methods = new CommonMethods();

                for (int i = 0; i < pointsFinalPunktir.Count - 1; ++i)
                {
                    DBPoint mp1point = (DBPoint)Trans.GetObject(pointsFinalPunktir[i], OpenMode.ForWrite);
                    DBPoint mp2point = (DBPoint)Trans.GetObject(pointsFinalPunktir[i + 1], OpenMode.ForWrite);// проблема
                    Line mpLine = new Line(new Point3d(mp1point.Position.X - 0.5, mp1point.Position.Y, 0), new Point3d(mp2point.Position.X + 0.5, mp2point.Position.Y, 0)); //рисуем основную линию междопутья
                    acBlkTblRec.AppendEntity(mpLine);
                    Trans.AddNewlyCreatedDBObject(mpLine, true);
                    allPoperEnts.Add(mpLine.ObjectId);
                    shrixiPutey.Add(mpLine.ObjectId);
                    //---рисуем маленькие косые линии-----
                    if (i == 0)
                    {
                        Point3d naklLine1point = new Point3d(mp1point.Position.X - 0.5, mp1point.Position.Y - 0.5, 0);
                        Point3d naklLine2point = new Point3d(mp1point.Position.X + 0.5, mp1point.Position.Y + 0.5, 0);
                        Line nakLine1 = new Line(naklLine1point, naklLine2point);
                        acBlkTblRec.AppendEntity(nakLine1);
                        Trans.AddNewlyCreatedDBObject(nakLine1, true);
                        shrixiPutey.Add(nakLine1.ObjectId);
                        allPoperEnts.Add(nakLine1.ObjectId);
                    }



                    Point3d naklLine3point = new Point3d(mp2point.Position.X - 0.5, mp2point.Position.Y - 0.5, 0);
                    Point3d naklLine4point = new Point3d(mp2point.Position.X + 0.5, mp2point.Position.Y + 0.5, 0);
                    Line nakLine2 = new Line(naklLine3point, naklLine4point);
                    acBlkTblRec.AppendEntity(nakLine2);
                    Trans.AddNewlyCreatedDBObject(nakLine2, true);
                    shrixiPutey.Add(nakLine2.ObjectId);
                    allPoperEnts.Add(nakLine2.ObjectId);
                    //-----ставим текст величины междопутья-------



                    double mpDist = methods.Vychisli_S(mp1point.Position, mp2point.Position) / horizScaleKoef;
                    String mpString = mpDist.ToString("0.00", CultureInfo.InvariantCulture);
                    DBText mpText = new DBText();
                    mpText.TextString = mpString;
                    mpText.Position = new Point3d(mp1point.Position.X + 0.5 * (mp2point.Position.X - mp1point.Position.X), mp1point.Position.Y, 0);
                    mpText.Height = Convert.ToDouble(form1.textBox3.Text);
                    acBlkTblRec.AppendEntity(mpText);
                    Trans.AddNewlyCreatedDBObject(mpText, true);
                    allPoperEnts.Add(mpText.ObjectId);
                    //далее передвигаем значение междопутья на середину, чтобы текст визуально был по центру отрезка
                    Point3d leftPt = mpText.GeometricExtents.MinPoint;
                    Point3d rightPt = mpText.GeometricExtents.MaxPoint;
                    double deltaX = 0.5 * (rightPt.X - leftPt.X);
                    mpText.Position = new Point3d(mpText.Position.X - deltaX, mpText.Position.Y + 0.3, mpText.Position.Z);
                    if (mpText.TextStyleName == "ATP")
                    {
                        mpText.Oblique = 0.2618;
                        mpText.WidthFactor = 0.8;
                    }

                    mezhputTEXTs.Add(mpText.ObjectId);


                }
            }
            //перемещаем текст с описанием пути наверх
            foreach (ObjectId opisId in colIdOpisanie)
            {
                DBText myOpisTxt = (DBText)Trans.GetObject(opisId, OpenMode.ForWrite);
                if ((myOpisTxt.TextString.Length > 5) && (myOpisTxt.TextString.StartsWith("путь")))
                {
                    DBPoint dBPointKonPunt = new DBPoint();
                    foreach (ObjectId pointID in pointsFinalPunktir)
                    {
                        DBPoint mp1point = (DBPoint)Trans.GetObject(pointID, OpenMode.ForWrite);
                        if (mp1point.Position.X == myOpisTxt.Position.X)
                        {
                            dBPointKonPunt.Position = mp1point.Position;
                        }
                    }
                    myOpisTxt.TextString = myOpisTxt.TextString.Substring(5);
                    myOpisTxt.Position = new Point3d(myOpisTxt.Position.X, dBPointKonPunt.Position.Y + 0.9, 0);
                    myOpisTxt.Rotation = 0;
                    //далее передвигаем значение на середину, чтобы текст визуально был по центру выноски
                    Point3d leftPtOp = myOpisTxt.GeometricExtents.MinPoint;
                    Point3d rightPtOp = myOpisTxt.GeometricExtents.MaxPoint;
                    double deltaXOp = 0.5 * (rightPtOp.X - leftPtOp.X);
                    myOpisTxt.Position = new Point3d(myOpisTxt.Position.X - deltaXOp, myOpisTxt.Position.Y, myOpisTxt.Position.Z);


                }
                else
                {
                    //передвигаем тексты с описаниями на 0.3 от линии, чтобы не накладывались
                    myOpisTxt.Position = new Point3d(myOpisTxt.Position.X - 0.3, myOpisTxt.Position.Y, myOpisTxt.Position.Z);

                }

            }
            //далее надо передвинуть вниз (на высоту текста+2.3) те тексты с расстояниями между точками (в самом низу), которые не помещаются в ячейку
            foreach (ObjectId DistTxtId in rasstTEXTs)
            {
                DBText myDistTxt = (DBText)Trans.GetObject(DistTxtId, OpenMode.ForWrite);//делаем АТР текст прямым, чтобы последующие выносные линии не смещались,
                                                                                         //а были ровно посередине
                if (myDistTxt.TextStyleName == "ATP")
                {
                    myDistTxt.Oblique = 0;
                    myDistTxt.WidthFactor = 1;
                }
                double ctrlDist = Convert.ToDouble(myDistTxt.TextString) * horizScaleKoef;
                double myDist = myDistTxt.GeometricExtents.MaxPoint.X - myDistTxt.GeometricExtents.MinPoint.X;
                myDistTxt.ColorIndex = 256;
                if ((myDist + 0.1) >= ctrlDist)
                {
                    myDistTxt.Position = new Point3d(myDistTxt.Position.X, myDistTxt.Position.Y - myDistTxt.Height - 2.3, myDistTxt.Position.Z);
                    //рисуем коротенькую палочку выносной линии
                    Polyline polylineVynos = new Polyline();
                    Point2d point_1 = new Point2d(myDistTxt.Position.X + 0.5 * (myDistTxt.GeometricExtents.MaxPoint.X - myDistTxt.GeometricExtents.MinPoint.X), vspomLine4.StartPoint.Y + 0.5 * (vspomLine3.StartPoint.Y - vspomLine4.StartPoint.Y));
                    Point2d point_2 = new Point2d(point_1.X, point_1.Y - 3.5);
                    polylineVynos.AddVertexAt(0, point_1, 0, 0, 0);
                    polylineVynos.AddVertexAt(1, point_2, 0, 0, 0);
                    acBlkTblRec.AppendEntity(polylineVynos);
                    Trans.AddNewlyCreatedDBObject(polylineVynos, true);
                    allPoperEnts.Add(polylineVynos.ObjectId);
                }
                if (myDistTxt.TextStyleName == "ATP")
                {
                    myDistTxt.Oblique = 0.2618;
                    myDistTxt.WidthFactor = 0.8;
                }
            }
            //далее создаем легенду поперечника
            Point2d startPointLegend = new Point2d(vspomLine4.EndPoint.X - 80, vspomLine4.StartPoint.Y);
            Point2d pointLegend_1 = new Point2d(startPointLegend.X, vspomLine2.StartPoint.Y);
            Point2d pointLegend_2 = new Point2d(startPointLegend.X, gorizontLine.StartPoint.Y);
            Point2d pointLegend_3 = new Point2d(startPointLegend.X + 30, gorizontLine.StartPoint.Y);
            Point2d pointLegend_4 = new Point2d(startPointLegend.X + 60, gorizontLine.StartPoint.Y);
            Point2d pointLegend_5 = new Point2d(startPointLegend.X + 60, vspomLine1.StartPoint.Y);
            Point2d pointLegend_6 = new Point2d(startPointLegend.X + 60, vspomLine2.StartPoint.Y);
            Point2d pointLegend_7 = new Point2d(startPointLegend.X + 60, vspomLine3.StartPoint.Y);
            Point2d pointLegend_8 = new Point2d(startPointLegend.X + 60, startPointLegend.Y);
            Point2d pointLegend_9 = new Point2d(startPointLegend.X + 30, startPointLegend.Y);
            Point2d pointLegend_10 = new Point2d(startPointLegend.X + 30, vspomLine3.StartPoint.Y);
            Point2d pointLegend_11 = new Point2d(startPointLegend.X + 30, vspomLine1.StartPoint.Y);

            Polyline polylineLegend_1 = new Polyline();
            polylineLegend_1.AddVertexAt(0, startPointLegend, 0, 0, 0);
            polylineLegend_1.AddVertexAt(1, pointLegend_2, 0, 0, 0);
            polylineLegend_1.AddVertexAt(2, pointLegend_4, 0, 0, 0);
            polylineLegend_1.AddVertexAt(3, pointLegend_8, 0, 0, 0);
            polylineLegend_1.Closed = true;
            acBlkTblRec.AppendEntity(polylineLegend_1);
            Trans.AddNewlyCreatedDBObject(polylineLegend_1, true);
            allPoperEnts.Add(polylineLegend_1.ObjectId);

            Polyline polylineLegend_2 = new Polyline();
            polylineLegend_2.AddVertexAt(0, pointLegend_1, 0, 0, 0);
            polylineLegend_2.AddVertexAt(1, pointLegend_6, 0, 0, 0);
            acBlkTblRec.AppendEntity(polylineLegend_2);
            Trans.AddNewlyCreatedDBObject(polylineLegend_2, true);
            allPoperEnts.Add(polylineLegend_2.ObjectId);

            Polyline polylineLegend_3 = new Polyline();
            polylineLegend_3.AddVertexAt(0, pointLegend_9, 0, 0, 0);
            polylineLegend_3.AddVertexAt(1, pointLegend_3, 0, 0, 0);
            acBlkTblRec.AppendEntity(polylineLegend_3);
            Trans.AddNewlyCreatedDBObject(polylineLegend_3, true);
            allPoperEnts.Add(polylineLegend_3.ObjectId);

            Polyline polylineLegend_4 = new Polyline();
            polylineLegend_4.AddVertexAt(0, pointLegend_11, 0, 0, 0);
            polylineLegend_4.AddVertexAt(1, pointLegend_5, 0, 0, 0);
            acBlkTblRec.AppendEntity(polylineLegend_4);
            Trans.AddNewlyCreatedDBObject(polylineLegend_4, true);
            allPoperEnts.Add(polylineLegend_4.ObjectId);

            Polyline polylineLegend_5 = new Polyline();
            polylineLegend_5.AddVertexAt(0, pointLegend_10, 0, 0, 0);
            polylineLegend_5.AddVertexAt(1, pointLegend_7, 0, 0, 0);
            acBlkTblRec.AppendEntity(polylineLegend_5);
            Trans.AddNewlyCreatedDBObject(polylineLegend_5, true);
            allPoperEnts.Add(polylineLegend_5.ObjectId);

            DBText textFakDan = new DBText();
            textFakDan.TextString = "ФАКТИЧЕСКИЕ";
            textFakDan.Height = Convert.ToDouble(form1.textBox4.Text);
            textFakDan.TextStyleId = CurTextstyle;
            if (textFakDan.TextStyleName == "ATP")
            {
                textFakDan.Oblique = 0.2618;
                textFakDan.WidthFactor = 0.8;
            }
            textFakDan.Position = new Point3d(startPointLegend.X + 2, startPointLegend.Y + 10, 0);
            acBlkTblRec.AppendEntity(textFakDan);
            Trans.AddNewlyCreatedDBObject(textFakDan, true);
            allPoperEnts.Add(textFakDan.ObjectId);

            DBText textDannye = new DBText();
            textDannye.TextString = "ДАННЫЕ";
            textDannye.Height = Convert.ToDouble(form1.textBox4.Text);
            textDannye.TextStyleId = CurTextstyle;
            if (textDannye.TextStyleName == "ATP")
            {
                textDannye.Oblique = 0.2618;
                textDannye.WidthFactor = 0.8;
            }
            textDannye.Position = new Point3d(startPointLegend.X + 5, startPointLegend.Y + 5, 0);
            acBlkTblRec.AppendEntity(textDannye);
            Trans.AddNewlyCreatedDBObject(textDannye, true);
            allPoperEnts.Add(textDannye.ObjectId);

            DBText textProekDan = new DBText();
            textProekDan.TextString = "ПРОЕКТНЫЕ";
            textProekDan.Height = Convert.ToDouble(form1.textBox4.Text);
            textProekDan.TextStyleId = CurTextstyle;
            if (textProekDan.TextStyleName == "ATP")
            {
                textProekDan.Oblique = 0.2618;
                textProekDan.WidthFactor = 0.8;
            }
            textProekDan.Position = new Point3d(pointLegend_1.X + 2, pointLegend_1.Y + 10, 0);
            acBlkTblRec.AppendEntity(textProekDan);
            Trans.AddNewlyCreatedDBObject(textProekDan, true);
            allPoperEnts.Add(textProekDan.ObjectId);

            DBText textDannye_2 = new DBText();
            textDannye_2.TextString = "ДАННЫЕ";
            textDannye_2.Height = Convert.ToDouble(form1.textBox4.Text);
            textDannye_2.TextStyleId = CurTextstyle;
            if (textDannye_2.TextStyleName == "ATP")
            {
                textDannye_2.Oblique = 0.2618;
                textDannye_2.WidthFactor = 0.8;
            }
            textDannye_2.Position = new Point3d(pointLegend_1.X + 5, pointLegend_1.Y + 5, 0);
            acBlkTblRec.AppendEntity(textDannye_2);
            Trans.AddNewlyCreatedDBObject(textDannye_2, true);
            allPoperEnts.Add(textDannye_2.ObjectId);

            DBText text_1 = new DBText();
            text_1.TextString = "РАССТОЯНИЕ, М";
            text_1.Height = Convert.ToDouble(form1.textBox4.Text);
            text_1.TextStyleId = CurTextstyle;
            text_1.WidthFactor = 0.8;
            if (text_1.TextStyleName == "ATP")
            {
                text_1.Oblique = 0.2618;
                text_1.WidthFactor = 0.8;
            }
            text_1.Position = new Point3d(pointLegend_9.X + 2, pointLegend_9.Y + 1.5, 0);
            acBlkTblRec.AppendEntity(text_1);
            Trans.AddNewlyCreatedDBObject(text_1, true);
            allPoperEnts.Add(text_1.ObjectId);

            DBText text_2 = new DBText();
            text_2.TextString = "РЕЛЬЕФА, М";
            text_2.Height = Convert.ToDouble(form1.textBox4.Text);
            text_2.TextStyleId = CurTextstyle;
            if (text_2.TextStyleName == "ATP")
            {
                text_2.Oblique = 0.2618;
                text_2.WidthFactor = 0.8;
            }
            text_2.Position = new Point3d(pointLegend_10.X + 5, pointLegend_10.Y + 4, 0);
            acBlkTblRec.AppendEntity(text_2);
            Trans.AddNewlyCreatedDBObject(text_2, true);
            allPoperEnts.Add(text_2.ObjectId);

            DBText text_3 = new DBText();
            text_3.TextString = "ОТМЕТКА";
            text_3.Height = Convert.ToDouble(form1.textBox4.Text);
            text_3.TextStyleId = CurTextstyle;
            if (text_3.TextStyleName == "ATP")
            {
                text_3.Oblique = 0.2618;
                text_3.WidthFactor = 0.8;
            }
            text_3.Position = new Point3d(pointLegend_10.X + 5, pointLegend_10.Y + 9, 0);
            acBlkTblRec.AppendEntity(text_3);
            Trans.AddNewlyCreatedDBObject(text_3, true);
            allPoperEnts.Add(text_3.ObjectId);

            DBText text_4 = new DBText();
            text_4.TextString = "РАССТОЯНИЕ, М";
            text_4.Height = Convert.ToDouble(form1.textBox4.Text);
            text_4.TextStyleId = CurTextstyle;
            text_4.WidthFactor = 0.8;
            if (text_4.TextStyleName == "ATP")
            {
                text_4.Oblique = 0.2618;
                text_4.WidthFactor = 0.8;
            }
            text_4.Position = new Point3d(pointLegend_11.X + 2, pointLegend_11.Y - 3.5, 0);
            acBlkTblRec.AppendEntity(text_4);
            Trans.AddNewlyCreatedDBObject(text_4, true);
            allPoperEnts.Add(text_4.ObjectId);

            DBText text_5 = new DBText();
            text_5.TextString = "ОТМЕТКА, М";
            text_5.Height = Convert.ToDouble(form1.textBox4.Text);
            text_5.TextStyleId = CurTextstyle;
            if (text_5.TextStyleName == "ATP")
            {
                text_5.Oblique = 0.2618;
                text_5.WidthFactor = 0.8;
            }
            text_5.Position = new Point3d(pointLegend_11.X + 5, pointLegend_11.Y + 5, 0);
            acBlkTblRec.AppendEntity(text_5);
            Trans.AddNewlyCreatedDBObject(text_5, true);
            allPoperEnts.Add(text_5.ObjectId);

            DBText text_6 = new DBText();
            string horMasht = form1.textBox1.Text;
            text_6.TextString = "М 1:" + horMasht + " ГОРИЗОНТАЛЬНЫЙ";
            text_6.Height = Convert.ToDouble(form1.textBox4.Text);
            text_6.TextStyleId = CurTextstyle;
            if (text_6.TextStyleName == "ATP")
            {
                text_6.Oblique = 0.2618;
                text_6.WidthFactor = 0.8;
            }
            text_6.Position = new Point3d(pointLegend_2.X, pointLegend_2.Y + 2, 0);
            acBlkTblRec.AppendEntity(text_6);
            Trans.AddNewlyCreatedDBObject(text_6, true);
            allPoperEnts.Add(text_6.ObjectId);

            DBText text_7 = new DBText();
            string vertMasht = form1.textBox2.Text;
            text_7.TextString = "М 1:" + vertMasht + " ВЕРТИКАЛЬНЫЙ";
            text_7.Height = Convert.ToDouble(form1.textBox4.Text);
            text_7.TextStyleId = CurTextstyle;
            if (text_7.TextStyleName == "ATP")
            {
                text_7.Oblique = 0.2618;
                text_7.WidthFactor = 0.8;
            }
            text_7.Position = new Point3d(pointLegend_2.X, pointLegend_2.Y + 6, 0);
            acBlkTblRec.AppendEntity(text_7);
            Trans.AddNewlyCreatedDBObject(text_7, true);
            allPoperEnts.Add(text_7.ObjectId);

            DBText text_8 = new DBText();

            text_8.TextString = poperCaption;
            text_8.Height = Convert.ToDouble(form1.textBox5.Text);
            text_8.TextStyleId = CurTextstyle;
            if (text_8.TextStyleName == "ATP")
            {
                text_8.Oblique = 0.2618;
                text_8.WidthFactor = 0.8;
            }
            text_8.Position = new Point3d(vspomLine4.StartPoint.X + 0.5 * (vspomLine4.EndPoint.X - vspomLine4.StartPoint.X), vspomLine4.StartPoint.Y - 8, 0);
            acBlkTblRec.AppendEntity(text_8);
            Trans.AddNewlyCreatedDBObject(text_8, true);
            allPoperEnts.Add(text_8.ObjectId);
            //далее передвигаем значение на середину, чтобы текст визуально был по центру поперечника
            Point3d leftPtZag = text_8.GeometricExtents.MinPoint;
            Point3d rightPtZag = text_8.GeometricExtents.MaxPoint;
            double deltaXZag = 0.5 * (rightPtZag.X - leftPtZag.X);
            text_8.Position = new Point3d(text_8.Position.X - deltaXZag, text_8.Position.Y, text_8.Position.Z);

            //теперь удаляем лишние, уже ненужные объекты(коллекции)
            foreach (ObjectId entityId in colIdCodes)
            {
                DBText textAny = (DBText)Trans.GetObject(entityId, OpenMode.ForWrite);
                textAny.Erase(true);
                textAny.Dispose();
            }
            colIdCodes.Clear();
            foreach (ObjectId entityId in colIdPodzemkaRasst)
            {
                DBText textAny = (DBText)Trans.GetObject(entityId, OpenMode.ForWrite);
                textAny.Erase(true);
                textAny.Dispose();
            }
            colIdPodzemkaRasst.Clear();
            foreach (ObjectId entityId in colIdRasst)
            {
                DBText textAny = (DBText)Trans.GetObject(entityId, OpenMode.ForWrite);
                textAny.Erase(true);
                textAny.Dispose();
            }
            colIdRasst.Clear();
            foreach (ObjectId entityId in colIdRasstIsso)
            {
                DBText textAny = (DBText)Trans.GetObject(entityId, OpenMode.ForWrite);
                textAny.Erase(true);
                textAny.Dispose();
            }
            colIdRasstIsso.Clear();

            foreach (ObjectId entityId in colIdPodzemkaAbsolut)
            {
                allPoperEnts.Add(entityId);
            }
            colIdPodzemkaAbsolut.Clear();

            foreach (ObjectId entityId in pointsFinalPunktir)
            {
                DBPoint pointAny = (DBPoint)Trans.GetObject(entityId, OpenMode.ForWrite);
                pointAny.Erase(true);
                pointAny.Dispose();
            }

            //вычисляем вектор для перемещения поперечника
            Vector3d myVector = new Point3d(startPointLegend.X, startPointLegend.Y, 0).GetVectorTo(startPoint3D);

            //перемещаем поперечник на первоначальную точку
            foreach (ObjectId entId in allPoperEnts)
            {
                Entity entAny = (Entity)Trans.GetObject(entId, OpenMode.ForWrite);
                entAny.TransformBy(Matrix3d.Displacement(myVector));
                // entAny.ColorIndex=123;
            }
        }


    }
}
