using System;
using System.Globalization;
using System.IO;

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
                                double deltaX = Round(Convert.ToDouble(znach[1]), 2);//меняем знак расстояния в зависимости от "лево" или "право"
                                if (deltaX == 0)
                                {
                                    changeZnak = false;
                                }
                                if (changeZnak == true)
                                {
                                    deltaX = -deltaX;
                                }
                                double deltaY = Round(Convert.ToDouble(znach[2]), 2);
                                if (Code == 0 || Code == 1 || Code == 501)
                                {
                                    minY = Min(minY, startPoint3D.Y + deltaY * vertScaleKoef);
                                    maxY = Max(maxY, startPoint3D.Y + deltaY * vertScaleKoef);
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
                                        String roundDist = Abs(deltaX).ToString("0.00", CultureInfo.InvariantCulture);
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
                                    newOtmetka.TextString = Convert.ToString(Round(deltaY, 2));
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
                                    newOtmetka.TextString = Convert.ToString(Round(deltaY, 2));
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
                                    newRasst.TextString = Convert.ToString(Round(deltaX, 2));

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
                                    newRasst.TextString = Convert.ToString(Round(deltaX, 2));

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
                    MyCommonFunctions.ChertiPoper(profLine, minY, maxY, colIdOpisanie, colIdOtmetki, colIdRasst, colIdCodes, Trans, acBlkTblRec, horizScaleKoef, vertScaleKoef, colIdPodzemkaRasst, colIdPodzemkaAbsolut, colIdRasstIsso, db.Textstyle, poperCaption, startPoint3D);
                    // ed.WriteMessage($"\nДлина коллекции с описаниями ={colIdOpisanie.Count}\n, длина массива с отметками = {colIdOtmetki.Count}\n, длина массива с расстояниями {colIdRasst.Count}\n, длина массива с кодами {colIdCodes.Count}\n");
                    Trans.Commit();
                    // docklock.Dispose();
                }
#if NCAD
                catch (Teigha.Runtime.Exception ex)
#else
                catch (Autodesk.AutoCAD.Runtime.Exception ex)

#endif
                {
                    ed.WriteMessage($"В процессе считывания данных поперечника возникла ошибка: \n{ex}");
                    Trans.Abort();
                }
                catch (System.Exception ex1)
                {
                    ed.WriteMessage($"В процессе считывания данных поперечника возникла ошибка: \n{ex1}");
                    Trans.Abort();
                }
            }



        }

    }
}
