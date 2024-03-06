using System;

#if NCAD
using HostMgd.ApplicationServices;
using HostMgd.EditorInput;
using Teigha.Colors;
using Teigha.DatabaseServices;
using Teigha.Geometry;
using Teigha.Runtime;
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
    public static class Sdelay_PoperCSCmd
    {
        [CommandMethod("Sdelay_PoperCS", CommandFlags.Modal | CommandFlags.UsePickSet | CommandFlags.Session)]
        // [Obsolete]
        public static void Sdelay_PoperCS()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            //ФУНКЦИЯ ДЛЯ СОЗДАНИЯ ПОПЕРЕЧНИКА ИЗ 3D-полилинии. 
            BlockTable acBlkTbl;   //объявляем переменные для базы с примитивами чертежа 
            BlockTableRecord acBlkTblRec;
            Double H1;
            //______создаем фильтр выбора линий поперечника____________________________________
            TypedValue[] tv = new TypedValue[1];
            tv.SetValue(new TypedValue((int)(DxfCode.Start), "POLYLINE"), 0); //проблема здесь
            SelectionFilter filter = new SelectionFilter(tv);
            PromptSelectionOptions pso = new PromptSelectionOptions();
            pso.MessageForAdding = "\nВыберите 3-d полилинии поперечников\n";
            pso.Keywords.Add("POLYLINE");
            pso.AllowSubSelections = true;
            pso.SingleOnly = false;
            pso.SinglePickInSpace = false;
            //_______________________________________________________________________________________________________________
            PromptSelectionResult result = ed.GetSelection(pso, filter); //команда пользователю выбрать на экране 3-d полилинии поперечников, выбор идёт в соответствии с фильтром и опциями
            if (result.Status == PromptStatus.OK) // в случае, если пользователь что-то выбрал, то далее идёт процесс создания поперечника
            {
                String myDWG = doc.Name; //считываем имя активного чертежа вместе с путём и расширением
                SelectionSet newSel = result.Value; // создаём переменную для записи в неё всех выбранных 3-д полилиний
                //___________создаем набор объектов автокада (характерные линии), пересечения с которыми будем искать____________________________
                TypedValue[] TvPoper = new TypedValue[2];
                TvPoper.SetValue(new TypedValue((int)(DxfCode.Start), "LWPOLYLINE,LINE,SPLINE"), 0);
                TvPoper.SetValue(new TypedValue((int)(DxfCode.LayerName), "1,2,3,4,5,6,7,8,10,12,13,14,20,23,24,28,29,Здания и строения,Инженерные сооружения,Путевое хозяйство,Автодорожное хозяйство,Объекты электропередачи,Гидрография,Откосы,Рельеф,Растительность,Ограждения,06_Инженерно-технические сооружения,03_Здания и строения,10_Границы покрытий и угодий,05_Элементы зданий,09_Путевое хозяйство,07_Объекты электропередачи,11_Гидрография,12_Рельеф,13_Растительность,14_Ограждения,Канализации,Дренажи,Водопроводы,Теплосети,Газопроводы"), 1);
                SelectionFilter filterPoper = new SelectionFilter(TvPoper);
                PromptSelectionResult resultPoper = ed.SelectAll(filterPoper);
                SelectionSet poperSel = resultPoper.Value;
                Point3dCollection SortCol = new Point3dCollection();
                //_________________________________________________________________________________________________________________________________
                foreach (SelectedObject sObj in newSel) //перебираем каждый выбранный объект (3д-полилинию)
                {
                    DocumentLock docklock = doc.LockDocument();
                    using (Transaction Trans = db.TransactionManager.StartTransaction()) // начинаем транзакцию
                    {
                        try //начинаем обработку с блоком конструкцией улавливания ошибок
                        {
                            acBlkTbl = (BlockTable)Trans.GetObject(db.BlockTableId, OpenMode.ForRead, false, true);      //открываем для чтения класс BlockTable
                            acBlkTblRec = Trans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false, true) as BlockTableRecord;

                            //_____________________с помощью транзакции добавляем в базу чертежа слой "подписи поперечников" (а если он там уже есть - то не добавляем)______________________________
                            LayerTable acLyrTbl = (LayerTable)Trans.GetObject(db.LayerTableId, OpenMode.ForRead);
                            String sLayerName = "подписи поперечников";

                            if (acLyrTbl.Has(sLayerName) == false)
                            {
                                LayerTableRecord acLyrTblRec = new LayerTableRecord();
                                // Устанавливаем слою нужный мне цвет по индексу 190 и ранее заданное имя
#if NCAD
                                acLyrTblRec.Color = Teigha.Colors.Color.FromColorIndex(ColorMethod.ByAci, 190);
#else
                                acLyrTblRec.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 190);
#endif
                                acLyrTblRec.Name = sLayerName;
                                // открываем таблицу слоев для записи
                                acLyrTbl.UpgradeOpen();
                                // записываем новый слой в таблицу слоев и в транзакцию
                                acLyrTbl.Add(acLyrTblRec);
                                Trans.AddNewlyCreatedDBObject(acLyrTblRec, true);
                            }
                            sLayerName = "коды поперечников";
                            if (acLyrTbl.Has(sLayerName) == false)
                            {
                                LayerTableRecord acLyrTblRec = new LayerTableRecord();
                                // Устанавливаем слою нужный мне цвет по индексу 191 и ранее заданное имя
#if NCAD
                                acLyrTblRec.Color = Teigha.Colors.Color.FromColorIndex(ColorMethod.ByAci, 191);
#else

                                acLyrTblRec.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 191);
#endif
                                acLyrTblRec.Name = sLayerName;
                                // открываем таблицу слоев для записи
                                acLyrTbl.UpgradeOpen();
                                // записываем новый слой в таблицу слоев и в транзакцию
                                acLyrTbl.Add(acLyrTblRec);
                                Trans.AddNewlyCreatedDBObject(acLyrTblRec, true);
                            }
                            //________________________________________________________________________________________________________________
                            Entity ent = Trans.GetObject(sObj.ObjectId, OpenMode.ForWrite) as Entity; //приводим выбранный объект к типу Entity
                            Polyline3d MyPl3d = Trans.GetObject(ent.ObjectId, OpenMode.ForWrite) as Polyline3d; //приводим выбранный объект от типа entity к нужному мне типу Polyline3d
                            Point3dCollection PointInVertCol = new Point3dCollection();
                            Point3dCollection GotovyPoper3dCol = new Point3dCollection();
                            Point3dCollection Points_GR = new Point3dCollection();
                            //-----если нажат соответствующий текстбокс, то ставим флажок и создаем заготовку под создание текстов междопутий
                            Form1 form1 = new Form1();
                            bool need_MP = false;
                            string textMP_style = "";
                            double textMP_height = 2;
                            string textMP_layer = "";
                            if (form1.checkBox1.Checked)
                            {
                                need_MP = true;

                                textMP_layer = form1.textBox10.Text;
                                textMP_style = form1.textBox12.Text;
                                textMP_height = Convert.ToDouble(form1.textBox11.Text);
                                if (acLyrTbl.Has(textMP_layer) == false)
                                {
                                    LayerTableRecord acLyrTblRec = new LayerTableRecord();
                                    // Устанавливаем слою нужный мне цвет по индексу 191 и ранее заданное имя
#if NCAD
                                    acLyrTblRec.Color = Teigha.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
#else
                                    acLyrTblRec.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
#endif
                                    acLyrTblRec.Name = textMP_layer;
                                    // открываем таблицу слоев для записи
                                    acLyrTbl.UpgradeOpen();
                                    // записываем новый слой в таблицу слоев и в транзакцию
                                    acLyrTbl.Add(acLyrTblRec);
                                    Trans.AddNewlyCreatedDBObject(acLyrTblRec, true);
                                }
                            }

                            //__________________процедура считывания координат из вершин 3-д полилинии____________________________________________________
                            foreach (ObjectId acObjIdVert in MyPl3d)                               //перебираем каждую вершину 3-д полилинии, это делается как ObjectId в 3-д полилинии
                            {
                                PolylineVertex3d Vert;   // объявляем переменную для каждой 3-д вершины
                                Vert = Trans.GetObject(acObjIdVert, OpenMode.ForRead) as PolylineVertex3d; //считываем вершину по её ObjectId
                                Point3d RoundMyPosition = new Point3d(Math.Round(Vert.Position.X, 2), Math.Round(Vert.Position.Y, 2), Math.Round(Vert.Position.Z, 3));
                                PointInVertCol.Add(RoundMyPosition);
                                GotovyPoper3dCol.Add(RoundMyPosition);
                            }
                            for (int i = 0; i <= PointInVertCol.Count - 2; i++)
                            {
                                //создаем отрезок из сегмента 3-д полилинии______________________________________________
                                Point3d Newstartpoint = new Point3d(PointInVertCol[i].X, PointInVertCol[i].Y, 0);
                                Point3d NewEndtpoint = new Point3d(PointInVertCol[i + 1].X, PointInVertCol[i + 1].Y, 0);
                                Line MyLine = new Line(Newstartpoint, NewEndtpoint);
                                //_______________________________________________________________________________________________
                                //перебираем каждую попавшуюся характерную линию
                                foreach (SelectedObject PopObj in poperSel)
                                {
                                    Entity popent = Trans.GetObject(PopObj.ObjectId, OpenMode.ForWrite, false, false) as Entity;
                                    //__________________________если пересекаемый объект - в слое 29 и фиолетового цвета - пропускаем его____________________________________
                                    if (popent.Layer == "29" && popent.ColorIndex == 6) continue;

                                    //______________________________________________________________________________________________________________
                                    Point3dCollection poper3dCol = new Point3dCollection();
                                    //находим точки пересечения сегмента полилинии разреза и характерной линии (может быть несколько)
                                    //  MyLine.IntersectWith(popent, Intersect.OnBothOperands, poper3dCol, 0, 0);
                                    MyLine.IntersectWith(popent, Intersect.OnBothOperands, poper3dCol, IntPtr.Zero, IntPtr.Zero);

                                    for (int j = 0; j <= poper3dCol.Count - 1; j++)
                                    {
                                        H1 = Vychisli_Z(PointInVertCol[i], poper3dCol[j], PointInVertCol[i + 1]);
                                        //_________________________________________создаем точку пересечения линии поперечника с характерной линией_________________________________________________________________________________
                                        Point3d PopPointWithH = new Point3d(Math.Round(poper3dCol[j].X, 2), Math.Round(poper3dCol[j].Y, 2), Math.Round(PointInVertCol[i].Z + H1, 3));
                                        //______________________________________Здесь вставляем модуль вставки в чертеж описания точки____________________________________________________________________________________

                                        SozdanieKodaTochkiPopera(acBlkTblRec, PopPointWithH, popent, MyLine);
                                        SozdaniePodpisiTochkiPopera(acBlkTblRec, PopPointWithH, popent, MyLine);
                                        //--если в настройках нажат соответствующий чекбокс, то создаем тексты с междопутьями на плане---------------
                                        if (need_MP && popent.Layer == "29")
                                        {
                                            Points_GR.Add(PopPointWithH);
                                        }
                                        //_____________________________________________________________________________________________________________________________
                                        //если характерная линия пересекает поперечник не в существующей точке линии перелома, то точку пересечения кидаем в новую коллекцию
                                        bool UslovieNalichiaTochki = ProveritNalichieTochki(GotovyPoper3dCol, PopPointWithH);
                                        if (UslovieNalichiaTochki == false)
                                        {
                                            GotovyPoper3dCol.Add(PopPointWithH);
                                        }
                                        //___________переходим к следующему пересечению сегмента и характерной линии____________________________
                                    }

                                }
                                //переходим к следующей характерной линии пересечения

                                MyLine.Dispose();
                            } //' переходим к следующему сегменту
                            SortCol = SortPoint3dCollection(GotovyPoper3dCol);
                            if (need_MP)
                            {

                                //Запускаем отдельную процедуру создания текстов междопутий
                                if (Points_GR.Count > 1)
                                {
                                    ed.WriteMessage($"На линии поперечника найдено путей:{Points_GR.Count}\n");
                                    // Point3dCollection SortPointsGRcol = SortPoint3dCollection(Points_GR);
                                    Point3dCollection ClearPointsGRcol = delete_dubles(Points_GR);
                                    ed.WriteMessage($"Из найденных в обработку междопутий ушло:{ClearPointsGRcol.Count}\n");

                                    SdelayMPtexts(ClearPointsGRcol, textMP_style, textMP_height, textMP_layer);
                                    ed.WriteMessage($"Функция \"SdelayMPtexts\" отработала успешно\n");
                                }
                                else
                                {
                                    ed.WriteMessage($"На линии поперечника найдено путей:{Points_GR.Count}\n");
                                }

                            }
                            else ed.WriteMessage("Процедура создания междопутий пропущена из-за настроек\n");


                            //очищаем коллекцию от дубликатов__________________________________
                            if (SortCol.Count > 2)
                            {

                                bool hasDubls = true;
                                while (hasDubls)
                                {
                                    int countDel = 0;
                                    for (int dublInd = 0; dublInd < SortCol.Count - 1; dublInd++)
                                    {

                                        //for (int k = 1;k< SortCol.Count;k++)
                                        //{
                                        double dubleDist = Vychisli_S(SortCol[dublInd], SortCol[dublInd + 1]);

                                        if (dubleDist < 0.01)
                                        {
                                            SortCol.RemoveAt(dublInd + 1);
                                            countDel++;
                                        }
                                        //}
                                    }
                                    if (countDel == 0)
                                    {
                                        hasDubls = false;
                                    }
                                }

                            }
                            //_________________________________________________________
                            Polyline3d FinalPop3dPline = new Polyline3d(Poly3dType.SimplePoly, SortCol, false);
                            FinalPop3dPline.SetPropertiesFrom(MyPl3d);
                            FinalPop3dPline.ColorIndex = 20;
                            FinalPop3dPline.LineWeight = LineWeight.LineWeight035;
                            acBlkTblRec.AppendEntity(FinalPop3dPline);
                            Trans.AddNewlyCreatedDBObject(FinalPop3dPline, true);

                            Trans.Commit(); //закрываем транзакцию с примитивами
                            docklock.Dispose();
                            GotovyPoper3dCol.Clear();
                            SortCol.Clear();
                            //finalCol.Clear();

                        }//конец блока try перед catch         
                        catch (System.Exception ex)
                        //в случае обнаружения ошибки пишем её описание и прерываем транзакцию

                        {
                            ed.WriteMessage("Ошибка " + ex.Message);
                            Trans.Abort();
                        }


                    }//end using 

                } //next

            }
            //в случае, если пользователь нажал клавишу ESC - выходим из команды
            ed.WriteMessage("\nНе выбраны 3-d полилинии поперов");
            return;


        }

        private static void SozdaniePodpisiTochkiPopera(BlockTableRecord acBlkTblRec, Point3d PopPointWithH, Entity popent, Line MyLine)
        {
            // BlockTable acBlkTbl;   //объявляем переменные для базы с примитивами чертежа 
            //BlockTableRecord acBlkTblRec;
            TypedValue[] TvOpisanie = new TypedValue[2];
            TvOpisanie.SetValue(new TypedValue((int)(DxfCode.Start), "TEXT"), 0);
            TvOpisanie.SetValue(new TypedValue((int)(DxfCode.LayerName), "подписи поперечников"), 1);
            SelectionFilter filterOpisanie = new SelectionFilter(TvOpisanie);
            PromptSelectionResult resultOpisanie = ed.SelectAll(filterOpisanie);
            SelectionSet poperOpisanie = resultOpisanie.Value;
            DBText TryOpisanie = new DBText();
            using (Transaction Trans1 = db.TransactionManager.StartTransaction())
            {
                // acBlkTbl = (BlockTable)Trans1.GetObject(db.BlockTableId, OpenMode.ForRead, false, true);      //открываем для чтения класс BlockTable
                //  acBlkTblRec = Trans1.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false, true) as BlockTableRecord;
                if (poperOpisanie is null)
                {
                    //return;
                    goto метка2;
                }
                foreach (SelectedObject ObjOpisanie in poperOpisanie)
                {
                    TryOpisanie = Trans1.GetObject(ObjOpisanie.ObjectId, OpenMode.ForWrite, false, false) as DBText; //System.InvalidOperationException: "Операция является недопустимой из-за текущего состояния объекта."

                    if (Math.Round(TryOpisanie.Position.X, 2) == PopPointWithH.X & Math.Round(TryOpisanie.Position.Y, 2) == PopPointWithH.Y)
                    {
                        return;
                    }  //End If
                } //Next
            метка2: String StrOpisanie = SdelayOpisanieTochkiPopera_3(popent, Trans1);
                DBText TxtOpisanie = new DBText();
                //With TxtOpisanie
                TxtOpisanie.TextString = StrOpisanie;
                TxtOpisanie.Position = PopPointWithH;
                TxtOpisanie.Layer = "подписи поперечников";
                TxtOpisanie.Justify = AttachmentPoint.BaseLeft;
                TxtOpisanie.Rotation = MyLine.Angle + 1.5708;
                TxtOpisanie.Height = 0.5;
                TxtOpisanie.ColorIndex = 190;
                TxtOpisanie.LineWeight = (LineWeight)0.3;
                //End With
                acBlkTblRec.AppendEntity(TxtOpisanie);
                Trans1.AddNewlyCreatedDBObject(TxtOpisanie, true);
                Trans1.Commit();


            }//end using
            TryOpisanie.Dispose();
        }//end sub

        private static Point3dCollection delete_dubles(Point3dCollection point3DCollection_with_doubles)
        {
            bool hasDubls = true;
            while (hasDubls)
            {
                int countDel = 0;
                for (int dublInd = 0; dublInd < point3DCollection_with_doubles.Count - 1; dublInd++)
                {

                    //for (int k = 1;k< SortCol.Count;k++)
                    //{
                    double dubleDist = Vychisli_S(point3DCollection_with_doubles[dublInd], point3DCollection_with_doubles[dublInd + 1]);

                    if (dubleDist < 0.01)
                    {
                        point3DCollection_with_doubles.RemoveAt(dublInd + 1);
                        countDel++;
                    }
                    //}
                }
                if (countDel == 0)
                {
                    hasDubls = false;
                }
            }
            return point3DCollection_with_doubles;
        }

        private static Point3dCollection SortPoint3dCollection(Point3dCollection GotovyPoper3dCol)
        {
            Point3dCollection SortCol = new Point3dCollection();
            int i, j;
            double S1;
            double[] ArrayOfDist = new double[GotovyPoper3dCol.Count];
            int[] ArrayOfIndex = new int[GotovyPoper3dCol.Count];
            using (Transaction Trans = db.TransactionManager.StartTransaction()) // начинаем транзакцию
            {
                try
                {
                    for (i = 0; i <= GotovyPoper3dCol.Count - 1; i++)
                    {
                        S1 = Vychisli_S(GotovyPoper3dCol[0], GotovyPoper3dCol[i]);
                        ArrayOfDist.SetValue(S1, i);
                        ArrayOfIndex.SetValue(i, i);
                    } //Next
                    Array.Sort(ArrayOfDist, ArrayOfIndex);
                    for (j = 0; j <= GotovyPoper3dCol.Count - 1; j++)
                    {
                        SortCol.Add(GotovyPoper3dCol[ArrayOfIndex[j]]);
                    } //Next


                }
                catch (System.Exception ex)
                {

                    ed.WriteMessage("\nЧто-то пошло не так...." + ex.Message);
                } //End Try 
            } //End Using
            return SortCol;
        } //End Function

        private static void SozdanieKodaTochkiPopera(BlockTableRecord acBlkTblRec, Point3d PopPointWithH, Entity popent, Line MyLine)
        {
            TypedValue[] TvKod = new TypedValue[2];
            TvKod.SetValue(new TypedValue((int)(DxfCode.Start), "TEXT"), 0);
            TvKod.SetValue(new TypedValue((int)(DxfCode.LayerName), "коды поперечников"), 1);
            SelectionFilter filterKod = new SelectionFilter(TvKod);
            PromptSelectionResult resultKod = ed.SelectAll(filterKod);
            SelectionSet poperKod = resultKod.Value;
            DBText TryKod = new DBText();
            using (Transaction Trans2 = db.TransactionManager.StartTransaction())

            {
                if (poperKod is null)
                {
                    goto метка3;
                }
                else
                {
                    foreach (SelectedObject ObjKod in poperKod)
                    {
                        TryKod = Trans2.GetObject(ObjKod.ObjectId, OpenMode.ForWrite, false, false) as DBText;
                        if ((Math.Round(TryKod.Position.X, 2) == PopPointWithH.X) & (Math.Round(TryKod.Position.Y, 2) == PopPointWithH.Y))
                        {
                            return;
                        }//end if
                    }                         //Next

                }   // End If                  

            метка3: TypedValue[] TvVspomPoint = new TypedValue[2];
                TvVspomPoint.SetValue(new TypedValue((int)(DxfCode.Start), "POINT"), 0);
                TvVspomPoint.SetValue(new TypedValue((int)(DxfCode.LayerName), "коды поперечников,подписи поперечников"), 1);
                SelectionFilter filterVspomPoint = new SelectionFilter(TvVspomPoint);
                PromptSelectionResult resultVspomPoint = ed.SelectAll(filterVspomPoint);
                SelectionSet poperVspomPoint = resultVspomPoint.Value;
                Point3dCollection poperVspomPointCol = new Point3dCollection();
                Boolean TryVspomPoint;
                if (poperVspomPoint is null)
                {
                    TryVspomPoint = false;
                }
                else
                {
                    foreach (SelectedObject VspomPointObj in poperVspomPoint)
                    {
                        DBPoint VspomPoint = Trans2.GetObject(VspomPointObj.ObjectId, OpenMode.ForWrite, false, false) as DBPoint;
                        poperVspomPointCol.Add(VspomPoint.Position);
                    } //Next
                    TryVspomPoint = ProveritNalichieTochki(poperVspomPointCol, PopPointWithH);
                } //End If
                String StrKod = SdelayKodTochkiPopera(popent, Trans2, TryVspomPoint);
                DBText AddKod = new DBText();
                AddKod.TextString = StrKod;
                AddKod.Position = PopPointWithH;
                AddKod.Layer = "коды поперечников";
                AddKod.Justify = AttachmentPoint.BaseLeft;
                AddKod.Rotation = MyLine.Angle + 1.5708 + 1.5708;
                AddKod.Height = 0.5;
                AddKod.ColorIndex = 210;
                AddKod.LineWeight = (LineWeight)0.3;
                acBlkTblRec.AppendEntity(AddKod);
                Trans2.AddNewlyCreatedDBObject(AddKod, true);
                Trans2.Commit();
            }//end using
            TryKod.Dispose();
        }//end sub

        private static bool ProveritNalichieTochki(Point3dCollection myPoint3dCol, Point3d myPoint3d)
        {
            bool myBool = new bool();
            foreach (Point3d Vertpoint3d in myPoint3dCol)
            {
                if (Vertpoint3d.X == myPoint3d.X & Vertpoint3d.Y == myPoint3d.Y)
                {
                    myBool = true;
                } //End If
                else
                {
                    myBool = false;
                }
            } //Next
            return myBool;
        } //End Function

        private static string SdelayKodTochkiPopera(Entity popent, Transaction Trans, bool TryVspomPoint)
        {
            string KodTochkiPopera = "0";
            //'в зависимости от того, с каким entity идет пересечение - создаем набор кодов по-умолчанию
            switch (popent.GetType().ToString())
            {
#if NCAD
                case "Teigha.DatabaseServices.Polyline":
#else
                case "Autodesk.AutoCAD.DatabaseServices.Polyline":
#endif
                    using (Polyline ProverkaPolyLine = Trans.GetObject(popent.ObjectId, OpenMode.ForRead) as Polyline)
                    {
                        switch (ProverkaPolyLine.Layer)
                        {
                            case "1":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;
                            case "3":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;
                            case "4":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;
                            case "5":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;
                            case "6":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;
                            case "7":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;
                            case "8":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;

                            case "10":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;
                            case "13":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;
                            case "14":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;
                            case "2":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "131";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "130";
                                        break;

                                }
                                break;
                            case "12":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "141";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "140";
                                        break;
                                }
                                break;
                            case "29":
                                KodTochkiPopera = "1";
                                break;
                            case "23":
                                switch (ProverkaPolyLine.Linetype)
                                {
                                    case "atp_472":
                                        KodTochkiPopera = "4";
                                        break;
                                    case "atp_473":
                                        KodTochkiPopera = "5";
                                        break;
                                    case "atp_474_1b":
                                        KodTochkiPopera = "6";
                                        break;
                                    case "atp_474_2a":
                                        KodTochkiPopera = "7";
                                        break;
                                    case "atp_475_1":
                                        KodTochkiPopera = "8";
                                        break;
                                    case "atp_476_3":
                                        KodTochkiPopera = "9";
                                        break;
                                }
                                break;
                            case "20":
                                switch (ProverkaPolyLine.Linetype)
                                {
                                    case "atp_280_1":
                                        KodTochkiPopera = "10";
                                        break;
                                }
                                break;
                        }
                        break;

                    } //end using
#if NCAD
                case "Teigha.DatabaseServices.Line":
#else
                case "Autodesk.AutoCAD.DatabaseServices.Line":
#endif
                    using (Line ProverkaPolyLine = Trans.GetObject(popent.ObjectId, OpenMode.ForRead) as Line)
                    {
                        switch (ProverkaPolyLine.Layer)
                        {
                            case "1":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;
                            case "3":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;
                            case "4":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;
                            case "5":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;
                            case "6":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;
                            case "7":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;
                            case "8":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;

                            case "10":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;
                            case "13":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;
                            case "14":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;
                            case "2":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "131";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "130";
                                        break;

                                }
                                break;
                            case "12":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "131";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "130";
                                        break;
                                }
                                break;
                            case "29":
                                KodTochkiPopera = "1";
                                break;
                            case "23":
                                switch (ProverkaPolyLine.Linetype)
                                {
                                    case "atp_472":
                                        KodTochkiPopera = "4";
                                        break;
                                    case "atp_473":
                                        KodTochkiPopera = "5";
                                        break;
                                    case "atp_474_1b":
                                        KodTochkiPopera = "6";
                                        break;
                                    case "atp_474_2a":
                                        KodTochkiPopera = "7";
                                        break;
                                    case "atp_475_1":
                                        KodTochkiPopera = "8";
                                        break;
                                    case "atp_476_3":
                                        KodTochkiPopera = "9";
                                        break;
                                }
                                break;
                            case "20":
                                switch (ProverkaPolyLine.Linetype)
                                {
                                    case "atp_280_1":
                                        KodTochkiPopera = "10";
                                        break;
                                }
                                break;
                        }
                        break;



                    }
#if NCAD
                case "Teigha.DatabaseServices.Spline":
#else
                case "Autodesk.AutoCAD.DatabaseServices.Spline":
#endif
                    using (Spline ProverkaPolyLine = Trans.GetObject(popent.ObjectId, OpenMode.ForRead) as Spline)
                    {
                        switch (ProverkaPolyLine.Layer)
                        {
                            case "1":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;
                            case "3":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;
                            case "4":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;
                            case "5":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;
                            case "6":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;
                            case "7":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;
                            case "8":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;

                            case "10":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;
                            case "13":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;
                            case "14":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "121";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "120";
                                        break;

                                }
                                break;
                            case "2":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "131";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "130";
                                        break;

                                }
                                break;
                            case "12":
                                switch (TryVspomPoint.ToString())
                                {
                                    case "True":
                                        KodTochkiPopera = "131";
                                        break;
                                    case "False":
                                        KodTochkiPopera = "130";
                                        break;
                                }
                                break;
                            case "29":
                                KodTochkiPopera = "1";
                                break;
                            case "23":
                                switch (ProverkaPolyLine.Linetype)
                                {
                                    case "atp_472":
                                        KodTochkiPopera = "4";
                                        break;
                                    case "atp_473":
                                        KodTochkiPopera = "5";
                                        break;
                                    case "atp_474_1b":
                                        KodTochkiPopera = "6";
                                        break;
                                    case "atp_474_2a":
                                        KodTochkiPopera = "7";
                                        break;
                                    case "atp_475_1":
                                        KodTochkiPopera = "8";
                                        break;
                                    case "atp_476_3":
                                        KodTochkiPopera = "9";
                                        break;
                                }
                                break;
                            case "20":
                                switch (ProverkaPolyLine.Linetype)
                                {
                                    case "atp_280_1":
                                        KodTochkiPopera = "10";
                                        break;
                                }
                                break;
                        }
                        break;
                    }
            }
            return KodTochkiPopera;
        }//end function


    }
}
