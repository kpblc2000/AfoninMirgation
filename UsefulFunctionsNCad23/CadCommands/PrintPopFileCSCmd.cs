using System;
using System.Globalization;
using Infrastructure;
using Useful_FunctionsCsh;

#if NCAD
using HostMgd.ApplicationServices;
using HostMgd.EditorInput;
using Teigha.Colors;
using Teigha.DatabaseServices;
using Teigha.Runtime;
using Teigha.Geometry;
#elif ACAD
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
#endif

namespace UsefulFunctionsNCad23.CadCommands
{
    public static class PrintPopFileCSCmd
    {
        [CommandMethod("PrintPopFileCS", CommandFlags.Modal | CommandFlags.UsePickSet)]
        public static void PrintPopFileCS()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            //ФУНКЦИЯ ДЛЯ создания файла поперечников .рор. 
            BlockTable acBlkTbl;   //объявляем переменные для базы с примитивами чертежа 
            BlockTableRecord acBlkTblRec;
            //______создаем фильтр выбора линии сечения как 3-d полилинии чертежа____________________________________
            TypedValue[] tv = new TypedValue[1];
            tv.SetValue(new TypedValue((int)(DxfCode.Start), "POLYLINE"), 0);
            SelectionFilter filter = new SelectionFilter(tv);
            PromptSelectionOptions pso = new PromptSelectionOptions();
            pso.MessageForAdding = "\nВыберите 3-d полилинии поперечников\n";
            pso.Keywords.Add("POLYLINE");
            pso.AllowSubSelections = true;
            pso.SingleOnly = true;
            pso.SinglePickInSpace = false;
            //_______________________________________________________________________________________________________________
            PromptSelectionResult result = ed.GetSelection(pso, filter); //команда пользователю выбрать на экране
                                                                         //3-d полилинии поперечника, выбор идёт в соответствии
                                                                         //с фильтром и опциями
            if (result.Status == PromptStatus.OK) // в случае, если пользователь что-то выбрал, то далее идёт процесс
                                                  // создания текстового файла
            {
                Form1 form1 = new Form1();
                SelectionSet newSel = result.Value; // создаём переменную для записи в неё всех выбранных 3-д полилиний
                string sInfo;  // задаем переменную для информационной строки будущего файла
                string code = "0";
                string Opisanie = "";
                double RasstDoZero;
                double Htochki;
                // double HforCode120 = -0.8;
                double HforCode120 = Convert.ToDouble(form1.textBox6.Text);
                // double HforCode130 = 3;
                double HforCode130 = Convert.ToDouble(form1.textBox8.Text);
                // double HforVX = 0.15;
                double HforVX = Convert.ToDouble(form1.textBox9.Text);
                // double HforTR = -1.70;
                double HforTR = Convert.ToDouble(form1.textBox7.Text);

                TypedValue[] CodeTV = new TypedValue[2];
                //задаём фильтр выбора объектов, находящихся в вершине (выбирать будем только тексты)
                CodeTV.SetValue(new TypedValue((int)DxfCode.Start, "TEXT"), 0);
                CodeTV.SetValue(new TypedValue((int)DxfCode.LayerName, "коды поперечников,подписи поперечников"), 1);
                SelectionFilter Codefilter = new SelectionFilter(CodeTV);

                PromptSelectionResult VertSelRes = ed.SelectAll(Codefilter);
                foreach (SelectedObject sObj in newSel)//перебираем каждый выбранный объект (3д-полилинию)
                {
                    using (Transaction Trans = db.TransactionManager.StartTransaction())
                    {
                        //_____________________с помощью транзакции добавляем в базу чертежа слой "расстояния до Zero" (а если он там уже есть - то не добавляем)______________________________
                        LayerTable acLyrTbl = (LayerTable)Trans.GetObject(db.LayerTableId, OpenMode.ForRead);
                        String sLayerName = "расстояния до Zero";

                        if (acLyrTbl.Has(sLayerName) == false)
                        {
                            LayerTableRecord acLyrTblRec = new LayerTableRecord();
                            // Устанавливаем слою нужный мне цвет по индексу 64 и ранее заданное имя
#if NCAD
                            acLyrTblRec.Color = Teigha.Colors.Color.FromColorIndex(ColorMethod.ByAci, 64);
#else
                            acLyrTblRec.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 64);
#endif
                            acLyrTblRec.Name = sLayerName;
                            // открываем таблицу слоев для записи
                            acLyrTbl.UpgradeOpen();
                            // записываем новый слой в таблицу слоев и в транзакцию
                            acLyrTbl.Add(acLyrTblRec);
                            Trans.AddNewlyCreatedDBObject(acLyrTblRec, true);

                        }

                        acBlkTbl = (BlockTable)Trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                        acBlkTblRec = Trans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;  //открываем пространство модели для записи
                        try
                        {
                            Entity ent = Trans.GetObject(sObj.ObjectId, OpenMode.ForWrite) as Entity; //'приводим выбранный объект к типу Entity
                            Polyline3d MyPl3d = Trans.GetObject(ent.ObjectId, OpenMode.ForWrite) as Polyline3d; //приводим выбранный объект от типа entity к нужному мне типу Polyline3d

                            PromptStringOptions popFilePromptStringOptions = new PromptStringOptions("\nВведите имя поперечника\n");
                            popFilePromptStringOptions.AllowSpaces = true;
                            PromptResult PopNameResult = ed.GetString(popFilePromptStringOptions);//Ввод имени поперечника в дальнейшем планируется
                                                                                                  //контролировать через userform

                            if (PopNameResult.Status != PromptStatus.OK)
                            {
                                return;
                            }
                            String PopName = PopNameResult.StringResult;
                            // String sFile = doc.Name + "_" + PopName + ".pop";
                            String sFile = doc.Name.Substring(0, doc.Name.Length - 4) + "_" + PopName + ".pop";
                            System.IO.StreamWriter writer = new System.IO.StreamWriter(sFile); //создаем текстовый файл для записи в нём данных
                            //печатаем заголовок .рор-файла________________________________
                            String Zagolovok = "/pop_" + PopName + "\n";
                            writer.Write(Zagolovok);
                            //________________________________________________________
                            //---циклом делаем запрос нулевой точки, если промахнулся и выбрал не то - выбирай заново
                            bool popal = false;
                            Point3d ZeroPoint = new Point3d();
                            do
                            {
                                PromptPointOptions ppo = new PromptPointOptions("\nУкажите нулевую точку на линии поперечника\n");
                                ppo.AllowNone = false;
                                PromptPointResult ZeroPointResult = ed.GetPoint(ppo);
                                ZeroPoint = ZeroPointResult.Value; //пользователь задает нулевую точку отсчета на поперечника (от нее "лево" и "право")
                                popal = point_is_vertex(ZeroPoint, MyPl3d, Trans);
                                if (popal == false)
                                {
                                    ed.WriteMessage("Вы не попали в вершину рабочей линии поперечника!\n");
                                }
                            } while (popal == false);

                            PolylineVertex3d Vert = new PolylineVertex3d();
                            CommonMethods methods = new CommonMethods();
                            foreach (ObjectId acObjIdVert in MyPl3d) //перебираем каждую вершину 3-д полилинии,
                                                                     //это делается как ObjectId в 3-д полилинии
                            {

                                Vert = (PolylineVertex3d)Trans.GetObject(acObjIdVert, OpenMode.ForRead); // объявляем переменную для каждой 3-д вершины, считывая вершину по её ObjectId в коллекции
                                if (VertSelRes.Status == PromptStatus.OK)
                                {
                                    SelectionSet CodeSS = VertSelRes.Value;
                                    foreach (SelectedObject CodeObj in CodeSS)//'если в рамку выбора попали объекты в соответствии с фильтром - то делаем их анализ
                                    {
                                        using (DBText CodeText = Trans.GetObject(CodeObj.ObjectId, OpenMode.ForWrite) as DBText)
                                        {
                                            switch (CodeText.Layer)
                                            {
                                                case "коды поперечников":
                                                    if (Vert.Position.X == CodeText.Position.X & Vert.Position.Y == CodeText.Position.Y)
                                                    {
                                                        code = CodeText.TextString;
                                                    }
                                                    break;
                                                case "подписи поперечников":
                                                    if (Vert.Position.X == CodeText.Position.X & Vert.Position.Y == CodeText.Position.Y)
                                                    {
                                                        Opisanie = CodeText.TextString;
                                                    }
                                                    break;
                                            }
                                        } //end using
                                    } //следующий объект с кодом
                                      //!!!!!!!!!!!!Надо сделать как сумма длин сегментов ломаной линии!!!!!!!!!!______________________________________________________
                                      //RasstDoZero = Math.Round(Vychisli_S(Vert.Position, ZeroPoint), 2);
                                    RasstDoZero = Math.Round(methods.Vychisli_LomDlinu(MyPl3d, Vert.Position, ZeroPoint), 2);
                                    //______________________________________________________________________________
                                    if (RasstDoZero == 0)
                                    {
                                        String Razdel = "******************************\n";
                                        writer.Write(Razdel);
                                    }
                                    //Dim Razdel As String = "******************************" & vbCrLf & code & "      " & RasstDoZero.ToString & "      " & Htochki.ToString & "      " & Opisanie & vbCrLf
                                    //определяем отметку в вершине, конкретизируя случаи для кодов 120,130
                                    Htochki = Math.Round(Vert.Position.Z, 2);
                                    switch (code)
                                    {
                                        case "120":
                                            Htochki = HforCode120;
                                            if (Opisanie == "воздухопровод")
                                            {
                                                Htochki = HforVX;
                                            }
                                            if (Opisanie.StartsWith("водопровод") == true || Opisanie.StartsWith("канализация") == true || Opisanie.StartsWith("теплотрасса") == true || Opisanie.StartsWith("газопровод") == true)
                                            {
                                                Htochki = HforTR;
                                            }
                                            break;
                                        case "130":
                                            Htochki = HforCode130;
                                            break;
                                        case "140":
                                            Htochki = HforCode130;
                                            break;
                                    }
                                    //____________________________________________________________________
                                    //запускаем процедуру печатания строки текстового файла поперечника
                                    sInfo = code + "      " + RasstDoZero.ToString() + "      " + Htochki.ToString() + "      " + Opisanie + "\n"; //записываем данные в строчку через разделитель пробелы
                                    writer.Write(sInfo); //записываем строку в файл
                                                         //_________________________________________________________________________________________________________
                                                         //----создаем в отдельном слое тексты с расстояниями до Zero-------------------

                                    DBText text_zero_rasst = new DBText();
                                    text_zero_rasst.Position = Vert.Position;
                                    //text_zero_rasst.TextString = Convert.ToString(RasstDoZero);
                                    text_zero_rasst.TextString = "L=" + RasstDoZero.ToString("0.00", CultureInfo.InvariantCulture);
                                    text_zero_rasst.Layer = "расстояния до Zero";
                                    text_zero_rasst.Height = 0.2;
                                    text_zero_rasst.Rotation = 0;
                                    text_zero_rasst.ColorIndex = 256;
                                    acBlkTblRec.AppendEntity(text_zero_rasst);
                                    Trans.AddNewlyCreatedDBObject(text_zero_rasst, true);
                                    //text_zero_rasst.Position = new Point3d(text_zero_rasst.Position.X, text_zero_rasst.Position.Y+0.05, 0);
                                    text_zero_rasst.Rotation = 4.7124;
                                    //---------------------------------------------------------------------------------------------------------------------------------


                                } //следующая вершина
                                else
                                {
                                    ed.WriteMessage("\nВ чертеже не найдено кодов и описаний поперечника \n");
                                    return;
                                }
                                code = "0";
                                Opisanie = "";
                            } //следующая вершина
                            sInfo = "*****************************       \n";
                            writer.Write(sInfo);
                            writer.Close(); //закрываем запись файла


                            ed.WriteMessage("\nСоздан файл " + sFile + "\n"); //вспомогательно пишем в редактор сообщение о созданном файле, путь к файлу оттуда можно будет посмотреть
                            Trans.Commit(); //закрываем транзакцию с примитивами
                        }
                        catch (System.Exception ex) //'в случае обнаружения ошибки пишем её описание и прерываем транзакцию
                        {
                            ed.WriteMessage("Ошибка " + ex.Message);
                            Trans.Abort();
                        } //End Try  
                    }//end using
                } //next переходим к следующей выбранной 3-д полилинии
            } // end if
            else //в случае, если пользователь нажал клавишу ESC - выходим из команды
            {
                ed.WriteMessage("\nНе выбраны 3-d полилинии путей\n");
                return;
            }
        } //end sub

        private static bool point_is_vertex(Point3d myPoint3D, Polyline3d myPolyline3D, Transaction Trans)
        {
            bool popal = false;

            foreach (ObjectId acObjIdVert in myPolyline3D) //перебираем каждую вершину 3-д полилинии,
                //это делается как ObjectId в 3-д полилинии
            {

                PolylineVertex3d Vert = (PolylineVertex3d)Trans.GetObject(acObjIdVert, OpenMode.ForRead);
                {
                    if ((Vert.Position.X == myPoint3D.X) && (Vert.Position.Y == myPoint3D.Y))
                    {
                        popal = true; break;
                    }
                }
            }
            return popal;
        }
    }
}
