using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Threading.Tasks;
using static System.Data.DataTable;
using System.Windows.Input;
using System.ComponentModel;
using static System.Math;
using Autodesk.Windows;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Internal;
using Autodesk.AutoCAD.Colors;
using System.Linq.Expressions;
using System.IO;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Drawing;
//using System.Activities.Expressions;

//[assembly: CommandClass(typeof(Useful_FunctionsCsh.AfoninCommands2))]


// This line is not mandatory, but improves loading performances

namespace Useful_FunctionsCsh
{
    public class Class1
    {
        Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
        Editor ed = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor;
        Database db = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;

        [CommandMethod("Sdelay_PoperCS", CommandFlags.Modal | CommandFlags.UsePickSet | CommandFlags.Session)]
        [Obsolete]
        public void Sdelay_PoperCS()
        {
            //ФУНКЦИЯ ДЛЯ СОЗДАНИЯ ПОПЕРЕЧНИКА ИЗ 3D-полилинии. 
            BlockTable acBlkTbl;   //объявляем переменные для базы с примитивами чертежа 
            BlockTableRecord acBlkTblRec;
            Double H1;
            //______создаем фильтр выбора линий поперечника____________________________________
            TypedValue[] tv = new TypedValue[1];
            tv.SetValue(new TypedValue((int)(DxfCode.Start), "POLYLINE"), 0); //проблема здесь
            SelectionFilter filter = new SelectionFilter(tv);
            PromptSelectionOptions pso = new PromptSelectionOptions();
            pso.MessageForAdding = "\nВыберите 3-d полилинии поперечников";
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
                TvPoper.SetValue(new TypedValue((int)(DxfCode.LayerName), "1,2,3,4,5,6,7,8,10,12,13,14,20,23,24,28,29"), 1);
                SelectionFilter filterPoper = new SelectionFilter(TvPoper);
                PromptSelectionResult resultPoper = ed.SelectAll(filterPoper);
                SelectionSet poperSel = resultPoper.Value;
                Point3dCollection  SortCol = new Point3dCollection();
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
                            LayerTable acLyrTbl = (LayerTable) Trans.GetObject(db.LayerTableId, OpenMode.ForRead);
                            String sLayerName = "подписи поперечников";

                            if (acLyrTbl.Has(sLayerName) == false)
                            {
                                LayerTableRecord acLyrTblRec = new LayerTableRecord();
                                // Устанавливаем слою нужный мне цвет по индексу 190 и ранее заданное имя
                                acLyrTblRec.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 190);
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
                                acLyrTblRec.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 191);
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
                            //__________________процедура считывания координат из вершин 3-д полилинии____________________________________________________
                            foreach (ObjectId acObjIdVert in MyPl3d)                               //перебираем каждую вершину 3-д полилинии, это делается как ObjectId в 3-д полилинии
                            {
                                PolylineVertex3d Vert;   // объявляем переменную для каждой 3-д вершины
                                Vert = Trans.GetObject(acObjIdVert, OpenMode.ForRead) as PolylineVertex3d; //считываем вершину по её ObjectId
                                Point3d RoundMyPosition = new Point3d(Math.Round(Vert.Position.X, 2), Math.Round(Vert.Position.Y, 2), Math.Round(Vert.Position.Z, 3));
                                PointInVertCol.Add(RoundMyPosition);
                                GotovyPoper3dCol.Add(RoundMyPosition);
                            }
                               for (int i=0; i<=PointInVertCol.Count - 2;i++)
                               {
                                    //создаем отрезок из сегмента 3-д полилинии______________________________________________
                                    Point3d Newstartpoint = new Point3d(PointInVertCol[i].X, PointInVertCol[i].Y, 0);
                                    Point3d NewEndtpoint = new Point3d(PointInVertCol[i + 1].X, PointInVertCol[i + 1].Y, 0);
                                    Line MyLine = new Line(Newstartpoint, NewEndtpoint);
                                    //_______________________________________________________________________________________________
                                    //перебираем каждую попавшуюся характерную линию
                                    foreach (SelectedObject PopObj in poperSel)
                                    {
                                       Entity  popent  = Trans.GetObject(PopObj.ObjectId, OpenMode.ForWrite, false, false) as Entity;
                                        //__________________________каждый выбранный элемент опускаем на отметку 0____________________________________


                                        //______________________________________________________________________________________________________________
                                        Point3dCollection poper3dCol = new Point3dCollection();
                                        //находим точки пересечения сегмента полилинии разреза и характерной линии (может быть несколько)
                                        MyLine.IntersectWith(popent, Intersect.OnBothOperands, poper3dCol, 0, 0);

                                        for (int j=0; j<= poper3dCol.Count - 1; j++)
                                        {
                                            H1 = Vychisli_Z(PointInVertCol[i], poper3dCol[j], PointInVertCol[i + 1]);
                                            //_________________________________________создаем точку пересечения линии поперечника с характерной линией_________________________________________________________________________________
                                            Point3d PopPointWithH = new Point3d(Math.Round(poper3dCol[j].X, 2), Math.Round(poper3dCol[j].Y, 2), Math.Round(PointInVertCol[i].Z + H1, 3));
                                            //______________________________________Здесь вставляем модуль вставки в чертеж описания точки____________________________________________________________________________________

                                            SozdanieKodaTochkiPopera(acBlkTblRec, PopPointWithH, popent, MyLine);
                                            SozdaniePodpisiTochkiPopera(acBlkTblRec, PopPointWithH, popent, MyLine);
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
                                 Polyline3d FinalPop3dPline = new Polyline3d (Poly3dType.SimplePoly, SortCol, false);  
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
                              catch (Autodesk.AutoCAD.Runtime.Exception ex)
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
        public double Vychisli_Z(Point3d p1, Point3d p2, Point3d p3)
        {
            Line line1 = new Line(new Point3d(p1.X, p1.Y, 0), new Point3d(p2.X, p2.Y, 0));
            Line line2 = new Line(new Point3d(p2.X, p2.Y, 0), new Point3d(p3.X, p3.Y, 0));
            Double S1 = line1.Length;
            Double S2 = line2.Length;
            Double Hfull = p3.Z - p1.Z;
            Double Spol = S1 + S2;
            Double TanA = Hfull / Spol;
                  Double H1 = S1 * TanA;
            line1.Dispose();
            line2.Dispose();
           return H1;
        }

        public void SozdanieKodaTochkiPopera(BlockTableRecord acBlkTblRec,Point3d PopPointWithH, Entity popent, Line MyLine)
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
        public bool ProveritNalichieTochki(Point3dCollection myPoint3dCol, Point3d myPoint3d)
        {
            bool myBool = new bool();
            foreach (Point3d Vertpoint3d in myPoint3dCol)
            {
            if (Vertpoint3d.X == myPoint3d.X & Vertpoint3d.Y == myPoint3d.Y)
                {
                    myBool= true;
                } //End If
                else
                {
                    myBool = false;
                }
            } //Next
            return myBool;
        } //End Function

        public bool ProveritNalichiePodpisiInTochka(Point3d myPoint3d)

        {
            bool myBool = new bool();
            TypedValue[] TvOpisanie = new TypedValue[2];
            TvOpisanie.SetValue(new TypedValue((int)(DxfCode.Start), "TEXT"), 0);
            TvOpisanie.SetValue(new TypedValue((int)(DxfCode.LayerName), "подписи поперечников"), 1);
            SelectionFilter filterOpisanie = new SelectionFilter(TvOpisanie);
            PromptSelectionResult resultOpisanie = ed.SelectAll(filterOpisanie);
            SelectionSet poperOpisanie = resultOpisanie.Value;
            DBText TryOpisanie = new DBText();
            if (poperOpisanie is null)
            {
                myBool = false;
            } //End If
            using (Transaction Trans = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject ObjOpisanie in poperOpisanie)
                {
                    TryOpisanie = Trans.GetObject(ObjOpisanie.ObjectId, OpenMode.ForWrite, false, false) as DBText;
                    if (TryOpisanie.Position.X == myPoint3d.X & TryOpisanie.Position.Y == myPoint3d.Y)
                    {
                        myBool = true;
                    } //End If
                    else
                    {
                        myBool = false;
                    }
                    
                } //Next
            } //End Using
            return myBool;
        } //End Function

        public Point3dCollection SortPoint3dCollection(Point3dCollection GotovyPoper3dCol)
        {
            Point3dCollection SortCol = new Point3dCollection();
            int i, j;
            double S1; 
           double[] ArrayOfDist=new double[GotovyPoper3dCol.Count];
            int[] ArrayOfIndex = new int[GotovyPoper3dCol.Count];
            using (Transaction Trans = db.TransactionManager.StartTransaction()) // начинаем транзакцию
            {
                try
                {    
                        for (i = 0; i<= GotovyPoper3dCol.Count - 1;i++)
                        {
                            S1 = Vychisli_S(GotovyPoper3dCol[0], GotovyPoper3dCol[i]);
                            ArrayOfDist.SetValue(S1, i);
                            ArrayOfIndex.SetValue(i, i);
                        } //Next
                    Array.Sort(ArrayOfDist, ArrayOfIndex);
                    for (j = 0; j<=GotovyPoper3dCol.Count - 1;j++)
                    {
                        SortCol.Add(GotovyPoper3dCol[ArrayOfIndex[j]]);
                    } //Next
                        
                    
                }
                catch(System.Exception ex)
                {

                    ed.WriteMessage("\nЧто-то пошло не так...." + ex.Message);
                } //End Try 
            } //End Using
            return SortCol;
        } //End Function

        public double Vychisli_S(Point3d p1, Point3d p2)
        {
            double S1 = 0;
            using (Line myLine = new Line(new Point3d(p1.X, p1.Y, 0), new Point3d(p2.X, p2.Y, 0)))
            {
                S1 = myLine.Length;
            }
            return S1;
        }// End Function
        public string SdelayKodTochkiPopera(Entity popent, Transaction Trans,bool TryVspomPoint)
        {
            string KodTochkiPopera = "0";
            //'в зависимости от того, с каким entity идет пересечение - создаем набор кодов по-умолчанию
            switch (popent.GetType().ToString())
            {
                case "Autodesk.AutoCAD.DatabaseServices.Polyline":
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
                    
                case "Autodesk.AutoCAD.DatabaseServices.Line":
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
                case "Autodesk.AutoCAD.DatabaseServices.Spline":
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
        public void SozdaniePodpisiTochkiPopera(BlockTableRecord acBlkTblRec, Point3d PopPointWithH, Entity popent, Line MyLine)
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
   метка2:          String StrOpisanie = SdelayOpisanieTochkiPopera_3(popent, Trans1);
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


        public String SdelayOpisanieTochkiPopera_3(Entity popent, Transaction Trans)
        {
            String OpisanieTochkiPopera = ProveritNalichiePodpisi(popent);
            //проверяем, не заданы ли заранее подписи. Если заданы - то должно сразу на Return,
            //если не заданы - то идет их создание "по-умолчанию"
            if (OpisanieTochkiPopera == "подпись не задана" || OpisanieTochkiPopera == "Подписей поперечников не существует") 
            {
                switch (popent.GetType().ToString())
                {
                    case "Autodesk.AutoCAD.DatabaseServices.Polyline":
                      using( Polyline ProverkaPolyLine = Trans.GetObject(popent.ObjectId, OpenMode.ForRead) as Polyline)
                        {
                            switch (ProverkaPolyLine.Linetype)
                            {
                                case "atp_122_v":
                                    OpisanieTochkiPopera = "Водопровод";
                                    break;
                                case "atp_121_v":
                                    OpisanieTochkiPopera = "Водопровод наземный";
                                    break;
                                case "atp_122_d":
                                    OpisanieTochkiPopera = "Дренаж";
                                    break;
                                case "atp_121_d":
                                    OpisanieTochkiPopera = "Дренаж наземный";
                                    break;
                                case "atp_122_k":
                                    OpisanieTochkiPopera = "Канализация";
                                    break;
                                case "atp_121_k":
                                    OpisanieTochkiPopera = "Канализация наземная";
                                    break;
                                case "atp_122_kn":
                                    OpisanieTochkiPopera = "Канализация напорная";
                                    break;
                                case "atp_122_kl":
                                    OpisanieTochkiPopera = "Канализация ливневая";
                                    break;
                                case "atp_122_g":
                                    OpisanieTochkiPopera = "Газопровод";
                                    break;
                                case "atp_121_g":
                                    OpisanieTochkiPopera = "Газопровод наземный";
                                    break;
                                case "atp_122_t":
                                    OpisanieTochkiPopera = "Теплопровод";
                                    break;
                                case "atp_121_t":
                                    OpisanieTochkiPopera = "Теплопровод наземный";
                                    break;
                                case "atp_119_3":
                                    switch (ProverkaPolyLine.Layer)
                                    {
                                        case "7":
                                            OpisanieTochkiPopera = "Каб.0.4кВ";
                                            break;
                                        case "10":
                                            OpisanieTochkiPopera = "Каб.СЦБ";
                                            break;
                                    } //End Select
                                    break;
                                case "atp_121_2":
                                    OpisanieTochkiPopera = "Каб.0.4кВ наземный";
                                    break;
                                case "atp_120_3b":
                                    OpisanieTochkiPopera = "Каб.0.4кВ подводный";
                                    break;
                                case "atp_119_1":
                                    OpisanieTochkiPopera = "Каб.10кВ";
                                    break;
                                case "atp_121_1":
                                    OpisanieTochkiPopera = "Каб.10кВ наземный";
                                    break;
                                case "atp_120_3a":
                                    OpisanieTochkiPopera = "Каб.10кВ подводный";
                                    break;
                                case "atp_133_t":
                                    OpisanieTochkiPopera = "Каб.связи";
                                    break;
                                case "atp_133":
                                    OpisanieTochkiPopera = "Телефонная канализация";
                                    break;
                                case "atp_121_3":
                                    OpisanieTochkiPopera = "Каб.связи наземный";
                                    break;
                                case "atp_122_vx":
                                    OpisanieTochkiPopera = "воздух.подземный";
                                    break;
                                case "atp_121_vx":
                                    OpisanieTochkiPopera = "воздухопровод";
                                    break;
                                case "atp_5x2":
                                    OpisanieTochkiPopera = "край грунтовой дороги";
                                    break;
                                case "atp_472":
                                    OpisanieTochkiPopera = "ж.б. забор";
                                    break;
                                case "atp_473":
                                    OpisanieTochkiPopera = "ж.б. ограждение";
                                    break;
                                case "atp_474_2a":
                                    OpisanieTochkiPopera = "мет.ограждение";
                                    break;
                                case "atp_474_1b":
                                    OpisanieTochkiPopera = "мет.забор";
                                    break;
                                case "atp_475_1":
                                    OpisanieTochkiPopera = "дер.забор";
                                    break;
                                case "atp_476_3":
                                    OpisanieTochkiPopera = "забор сетка-рабица";
                                    break;
                                case "atp_476_1":
                                    OpisanieTochkiPopera = "колюч.проволока";
                                    break;
                                case "atp_477":
                                    OpisanieTochkiPopera = "деревян.плетень";
                                    break;
                                case "atp_476_2a":
                                    OpisanieTochkiPopera = "гладкая проволока";
                                    break;
                                case "atp_475_4a":
                                    OpisanieTochkiPopera = "деревян.с капитал. опорами";
                                    break;
                                case "atp_475_2":
                                    OpisanieTochkiPopera = "деревян.решетчатый";
                                    break;
                                case "atp_278_6b":
                                    OpisanieTochkiPopera = "метал.парапет";
                                    break;
                                case "atp_278_5":
                                    OpisanieTochkiPopera = "камен.парапет";
                                    break;
                                case "atp_1.5x1.5":
                                    switch (ProverkaPolyLine.Layer)
                                    {
                                        case "20":
                                            OpisanieTochkiPopera = "Край отмостки";
                                            break;
                                        case "23":
                                        OpisanieTochkiPopera = "край асф.";
                                            break;
                                    }
                                    break;

                                case "Continuous":
                                switch (ProverkaPolyLine.Layer)
                                    {
                                        case "20":
                                            OpisanieTochkiPopera = "Стена здания";
                                            break;
                                        case "23":
                                            OpisanieTochkiPopera = "край дороги";
                                            break;
                                    }

                                    break;


                            }
                            switch (ProverkaPolyLine.Layer)
                            {
                                case "2":
                                switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case 7:
                                            OpisanieTochkiPopera = "ЛЭП";
                                            break;
                                    }
                                    break;

                                case "12":
                                switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case int i when i != 7:
                                            OpisanieTochkiPopera = "линия связи";
                                            break;
                                    }
                                    break;
                                case "14":
                                    OpisanieTochkiPopera = "Трубопровод спец.назначения";
                                    break;
                                case "23":
                                switch (ProverkaPolyLine.Linetype)
                                    {
                                        case "atp_0x1":
                                            OpisanieTochkiPopera = "граница угодий";
                                            break;
                                        case "atp_0.2x1.5":
                                            OpisanieTochkiPopera = "граница угодий";
                                            break;
                                    }
                                switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case 3:
                                            OpisanieTochkiPopera = "край канавы";
                                            break;
                                    }
                                    break;
                                case "28":
                                switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case int i when i != 7:
                                            OpisanieTochkiPopera = "урез воды";
                                            break;
                                    }
                                    break;
                                case "24":
                                switch (ProverkaPolyLine.Linetype)
                                    {
                                        case "atp_0x1":
                                            OpisanieTochkiPopera = "край деревьев";
                                            break;
                                        case "atp_0.2x1.5":
                                            OpisanieTochkiPopera = "край деревьев";
                                            break;
                                    }
                                    break;
                                case "29":
                                switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case int i when i != 6:
                                            OpisanieTochkiPopera = "путь";
                                            break;
                                    }
                                    break;
                            }

                        } //end using

                        break;
                    case "Autodesk.AutoCAD.DatabaseServices.Line":
                        using (Line ProverkaPolyLine = Trans.GetObject(popent.ObjectId, OpenMode.ForRead) as Line)
                        {
                            switch (ProverkaPolyLine.Linetype)
                            {
                                case "atp_122_v":
                                    OpisanieTochkiPopera = "Водопровод";
                                    break;
                                case "atp_121_v":
                                    OpisanieTochkiPopera = "Водопровод наземный";
                                    break;
                                case "atp_122_d":
                                    OpisanieTochkiPopera = "Дренаж";
                                    break;
                                case "atp_121_d":
                                    OpisanieTochkiPopera = "Дренаж наземный";
                                    break;
                                case "atp_122_k":
                                    OpisanieTochkiPopera = "Канализация";
                                    break;
                                case "atp_121_k":
                                    OpisanieTochkiPopera = "Канализация наземная";
                                    break;
                                case "atp_122_kn":
                                    OpisanieTochkiPopera = "Канализация напорная";
                                    break;
                                case "atp_122_kl":
                                    OpisanieTochkiPopera = "Канализация ливневая";
                                    break;
                                case "atp_122_g":
                                    OpisanieTochkiPopera = "Газопровод";
                                    break;
                                case "atp_121_g":
                                    OpisanieTochkiPopera = "Газопровод наземный";
                                    break;
                                case "atp_122_t":
                                    OpisanieTochkiPopera = "Теплопровод";
                                    break;
                                case "atp_121_t":
                                    OpisanieTochkiPopera = "Теплопровод наземный";
                                    break;
                                case "atp_119_3":
                                    switch (ProverkaPolyLine.Layer)
                                    {
                                        case "7":
                                            OpisanieTochkiPopera = "Каб.0.4кВ";
                                            break;
                                        case "10":
                                            OpisanieTochkiPopera = "Каб.СЦБ";
                                            break;
                                    } //End Select
                                    break;
                                case "atp_121_2":
                                    OpisanieTochkiPopera = "Каб.0.4кВ наземный";
                                    break;
                                case "atp_120_3b":
                                    OpisanieTochkiPopera = "Каб.0.4кВ подводный";
                                    break;
                                case "atp_119_1":
                                    OpisanieTochkiPopera = "Каб.10кВ";
                                    break;
                                case "atp_121_1":
                                    OpisanieTochkiPopera = "Каб.10кВ наземный";
                                    break;
                                case "atp_120_3a":
                                    OpisanieTochkiPopera = "Каб.10кВ подводный";
                                    break;
                                case "atp_133_t":
                                    OpisanieTochkiPopera = "Каб.связи";
                                    break;
                                case "atp_133":
                                    OpisanieTochkiPopera = "Телефонная канализация";
                                    break;
                                case "atp_121_3":
                                    OpisanieTochkiPopera = "Каб.связи наземный";
                                    break;
                                case "atp_122_vx":
                                    OpisanieTochkiPopera = "воздух.подземный";
                                    break;
                                case "atp_121_vx":
                                    OpisanieTochkiPopera = "воздухопровод";
                                    break;
                                case "atp_5x2":
                                    OpisanieTochkiPopera = "край грунтовой дороги";
                                    break;
                                case "atp_472":
                                    OpisanieTochkiPopera = "ж.б. забор";
                                    break;
                                case "atp_473":
                                    OpisanieTochkiPopera = "ж.б. ограждение";
                                    break;
                                case "atp_474_2a":
                                    OpisanieTochkiPopera = "мет.ограждение";
                                    break;
                                case "atp_474_1b":
                                    OpisanieTochkiPopera = "мет.забор";
                                    break;
                                case "atp_475_1":
                                    OpisanieTochkiPopera = "дер.забор";
                                    break;
                                case "atp_476_3":
                                    OpisanieTochkiPopera = "забор сетка-рабица";
                                    break;
                                case "atp_476_1":
                                    OpisanieTochkiPopera = "колюч.проволока";
                                    break;
                                case "atp_477":
                                    OpisanieTochkiPopera = "деревян.плетень";
                                    break;
                                case "atp_476_2a":
                                    OpisanieTochkiPopera = "гладкая проволока";
                                    break;
                                case "atp_475_4a":
                                    OpisanieTochkiPopera = "деревян.с капитал. опорами";
                                    break;
                                case "atp_475_2":
                                    OpisanieTochkiPopera = "деревян.решетчатый";
                                    break;
                                case "atp_278_6b":
                                    OpisanieTochkiPopera = "метал.парапет";
                                    break;
                                case "atp_278_5":
                                    OpisanieTochkiPopera = "камен.парапет";
                                    break;
                                case "atp_1.5x1.5":
                                    switch (ProverkaPolyLine.Layer)
                                    {
                                        case "20":
                                            OpisanieTochkiPopera = "Край отмостки";
                                            break;
                                        case "23":
                                            OpisanieTochkiPopera = "край асф.";
                                            break;
                                    }
                                    break;

                                case "Continuous":
                                    switch (ProverkaPolyLine.Layer)
                                    {
                                        case "20":
                                            OpisanieTochkiPopera = "Стена здания";
                                            break;
                                        case "23":
                                            OpisanieTochkiPopera = "край дороги";
                                            break;
                                    }

                                    break;


                            }
                            switch (ProverkaPolyLine.Layer)
                            {
                                case "2":
                                    switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case 7:
                                            OpisanieTochkiPopera = "ЛЭП";
                                            break;
                                    }
                                    break;

                                case "12":
                                    switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case int i when i != 7:
                                            OpisanieTochkiPopera = "линия связи";
                                            break;
                                    }
                                    break;
                                case "14":
                                    OpisanieTochkiPopera = "Трубопровод спец.назначения";
                                    break;
                                case "23":
                                    switch (ProverkaPolyLine.Linetype)
                                    {
                                        case "atp_0x1":
                                            OpisanieTochkiPopera = "граница угодий";
                                            break;
                                        case "atp_0.2x1.5":
                                            OpisanieTochkiPopera = "граница угодий";
                                            break;
                                    }
                                    switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case 3:
                                            OpisanieTochkiPopera = "край канавы";
                                            break;
                                    }
                                    break;
                                case "28":
                                    switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case int i when i != 7:
                                            OpisanieTochkiPopera = "урез воды";
                                            break;
                                    }
                                    break;
                                case "24":
                                    switch (ProverkaPolyLine.Linetype)
                                    {
                                        case "atp_0x1":
                                            OpisanieTochkiPopera = "край деревьев";
                                            break;
                                        case "atp_0.2x1.5":
                                            OpisanieTochkiPopera = "край деревьев";
                                            break;
                                    }
                                    break;
                                case "29":
                                    switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case int i when i != 6:
                                            OpisanieTochkiPopera = "путь";
                                            break;
                                    }
                                    break;
                            }

                        } //end using

                        break;
                    case "Autodesk.AutoCAD.DatabaseServices.Spline":
                        using (Spline ProverkaPolyLine = Trans.GetObject(popent.ObjectId, OpenMode.ForRead) as Spline)
                        {
                            switch (ProverkaPolyLine.Linetype)
                            {
                                case "atp_122_v":
                                    OpisanieTochkiPopera = "Водопровод";
                                    break;
                                case "atp_121_v":
                                    OpisanieTochkiPopera = "Водопровод наземный";
                                    break;
                                case "atp_122_d":
                                    OpisanieTochkiPopera = "Дренаж";
                                    break;
                                case "atp_121_d":
                                    OpisanieTochkiPopera = "Дренаж наземный";
                                    break;
                                case "atp_122_k":
                                    OpisanieTochkiPopera = "Канализация";
                                    break;
                                case "atp_121_k":
                                    OpisanieTochkiPopera = "Канализация наземная";
                                    break;
                                case "atp_122_kn":
                                    OpisanieTochkiPopera = "Канализация напорная";
                                    break;
                                case "atp_122_kl":
                                    OpisanieTochkiPopera = "Канализация ливневая";
                                    break;
                                case "atp_122_g":
                                    OpisanieTochkiPopera = "Газопровод";
                                    break;
                                case "atp_121_g":
                                    OpisanieTochkiPopera = "Газопровод наземный";
                                    break;
                                case "atp_122_t":
                                    OpisanieTochkiPopera = "Теплопровод";
                                    break;
                                case "atp_121_t":
                                    OpisanieTochkiPopera = "Теплопровод наземный";
                                    break;
                                case "atp_119_3":
                                    switch (ProverkaPolyLine.Layer)
                                    {
                                        case "7":
                                            OpisanieTochkiPopera = "Каб.0.4кВ";
                                            break;
                                        case "10":
                                            OpisanieTochkiPopera = "Каб.СЦБ";
                                            break;
                                    } //End Select
                                    break;
                                case "atp_121_2":
                                    OpisanieTochkiPopera = "Каб.0.4кВ наземный";
                                    break;
                                case "atp_120_3b":
                                    OpisanieTochkiPopera = "Каб.0.4кВ подводный";
                                    break;
                                case "atp_119_1":
                                    OpisanieTochkiPopera = "Каб.10кВ";
                                    break;
                                case "atp_121_1":
                                    OpisanieTochkiPopera = "Каб.10кВ наземный";
                                    break;
                                case "atp_120_3a":
                                    OpisanieTochkiPopera = "Каб.10кВ подводный";
                                    break;
                                case "atp_133_t":
                                    OpisanieTochkiPopera = "Каб.связи";
                                    break;
                                case "atp_133":
                                    OpisanieTochkiPopera = "Телефонная канализация";
                                    break;
                                case "atp_121_3":
                                    OpisanieTochkiPopera = "Каб.связи наземный";
                                    break;
                                case "atp_122_vx":
                                    OpisanieTochkiPopera = "воздух.подземный";
                                    break;
                                case "atp_121_vx":
                                    OpisanieTochkiPopera = "воздухопровод";
                                    break;
                                case "atp_5x2":
                                    OpisanieTochkiPopera = "край грунтовой дороги";
                                    break;
                                case "atp_472":
                                    OpisanieTochkiPopera = "ж.б. забор";
                                    break;
                                case "atp_473":
                                    OpisanieTochkiPopera = "ж.б. ограждение";
                                    break;
                                case "atp_474_2a":
                                    OpisanieTochkiPopera = "мет.ограждение";
                                    break;
                                case "atp_474_1b":
                                    OpisanieTochkiPopera = "мет.забор";
                                    break;
                                case "atp_475_1":
                                    OpisanieTochkiPopera = "дер.забор";
                                    break;
                                case "atp_476_3":
                                    OpisanieTochkiPopera = "забор сетка-рабица";
                                    break;
                                case "atp_476_1":
                                    OpisanieTochkiPopera = "колюч.проволока";
                                    break;
                                case "atp_477":
                                    OpisanieTochkiPopera = "деревян.плетень";
                                    break;
                                case "atp_476_2a":
                                    OpisanieTochkiPopera = "гладкая проволока";
                                    break;
                                case "atp_475_4a":
                                    OpisanieTochkiPopera = "деревян.с капитал. опорами";
                                    break;
                                case "atp_475_2":
                                    OpisanieTochkiPopera = "деревян.решетчатый";
                                    break;
                                case "atp_278_6b":
                                    OpisanieTochkiPopera = "метал.парапет";
                                    break;
                                case "atp_278_5":
                                    OpisanieTochkiPopera = "камен.парапет";
                                    break;
                                case "atp_1.5x1.5":
                                    switch (ProverkaPolyLine.Layer)
                                    {
                                        case "20":
                                            OpisanieTochkiPopera = "Край отмостки";
                                            break;
                                        case "23":
                                            OpisanieTochkiPopera = "край асф.";
                                            break;
                                    }
                                    break;

                                case "Continuous":
                                    switch (ProverkaPolyLine.Layer)
                                    {
                                        case "20":
                                            OpisanieTochkiPopera = "Стена здания";
                                            break;
                                        case "23":
                                            OpisanieTochkiPopera = "край дороги";
                                            break;
                                    }

                                    break;


                            }
                            switch (ProverkaPolyLine.Layer)
                            {
                                case "2":
                                    switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case 7:
                                            OpisanieTochkiPopera = "ЛЭП";
                                            break;
                                    }
                                    break;

                                case "12":
                                    switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case int i when i != 7:
                                            OpisanieTochkiPopera = "линия связи";
                                            break;
                                    }
                                    break;
                                case "14":
                                    OpisanieTochkiPopera = "Трубопровод спец.назначения";
                                    break;
                                case "23":
                                    switch (ProverkaPolyLine.Linetype)
                                    {
                                        case "atp_0x1":
                                            OpisanieTochkiPopera = "граница угодий";
                                            break;
                                        case "atp_0.2x1.5":
                                            OpisanieTochkiPopera = "граница угодий";
                                            break;
                                    }
                                    switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case 3:
                                            OpisanieTochkiPopera = "край канавы";
                                            break;
                                    }
                                    break;
                                case "28":
                                    switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case int i when i != 7:
                                            OpisanieTochkiPopera = "урез воды";
                                            break;
                                    }
                                    break;
                                case "24":
                                    switch (ProverkaPolyLine.Linetype)
                                    {
                                        case "atp_0x1":
                                            OpisanieTochkiPopera = "край деревьев";
                                            break;
                                        case "atp_0.2x1.5":
                                            OpisanieTochkiPopera = "край деревьев";
                                            break;
                                    }
                                    break;
                                case "29":
                                    switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case int i when i != 6:
                                            OpisanieTochkiPopera = "путь";
                                            break;
                                    }
                                    break;
                            }

                        } //end using

                        break;
                }



            }
            //в зависимости от того, с каким entity идет пересечение - создаем набор кодов по-умолчанию


            return OpisanieTochkiPopera;

        } //end function
        public String ProveritNalichiePodpisi(Entity popent)
        {
            String PodpisString = "подпись не задана";
            //задаем слой, в котором будут располагаться подписи поперечников ("подписи поперечников").
            //Если в чертеже существует текст в данном слое, точка вставки которого эквивалентна начальной точке
            //пересекаемой полилинии, то содержимое текста передается в подпись.
            //Если такого текста нет, то передается условная строка "подпись не задана"
            TypedValue[] TvPodpis = new TypedValue[2];
            TvPodpis.SetValue(new TypedValue((int)(DxfCode.Start), "TEXT"), 0);
            TvPodpis.SetValue(new TypedValue((int)(DxfCode.LayerName), "подписи поперечников"), 1);
            SelectionFilter filterPodpis = new SelectionFilter(TvPodpis);
            PromptSelectionResult resultPodpis = ed.SelectAll(filterPodpis);
            SelectionSet PodpisSel = resultPodpis.Value;
            if (PodpisSel is null)
            {
                PodpisString = "Подписей поперечников не существует";
                return PodpisString;
                //Exit Function
            } //End If
            using (Transaction Trans = db.TransactionManager.StartTransaction())
            {
                try
                {
                    switch (popent.GetType().ToString())
                    {
                        case "Autodesk.AutoCAD.DatabaseServices.Line":
                            using (Line ProverkaLine = Trans.GetObject(popent.ObjectId, OpenMode.ForRead) as Line)
                            {
                                foreach (SelectedObject PodpisObj in PodpisSel)
                                {
                                    DBText ProverkaText = Trans.GetObject(PodpisObj.ObjectId, OpenMode.ForRead) as DBText;
                                    if (ProverkaText.Position.X == ProverkaLine.StartPoint.X & ProverkaText.Position.Y == ProverkaLine.StartPoint.Y)
                                    {    
                                        PodpisString = ProverkaText.TextString;
                                    } //End If
                                    ProverkaText.Dispose();
                                } //Next
                            } //end using
                            break;
                        case "Autodesk.AutoCAD.DatabaseServices.Polyline":
                            using (Polyline ProverkaPolyLine = Trans.GetObject(popent.ObjectId, OpenMode.ForRead) as Polyline)
                            {
                                foreach (SelectedObject PodpisObj in PodpisSel)
                                {
                                    DBText ProverkaText = Trans.GetObject(PodpisObj.ObjectId, OpenMode.ForRead) as DBText;
                                    if (ProverkaText.Position.X == ProverkaPolyLine.StartPoint.X & ProverkaText.Position.Y == ProverkaPolyLine.StartPoint.Y)
                                    {
                                        PodpisString = ProverkaText.TextString;
                                    } //End If
                                    ProverkaText.Dispose();
                                } //Next
                            } //end using
                            break;
                    } //End Select
                }
                 catch (Autodesk.AutoCAD.Runtime.Exception ex )
                {
                    PodpisString = "В процессе поиска подписи поперечника произошла ошибка: " + ex.Message;
                    ed.WriteMessage(PodpisString);
                } //End Try
            }//end using
                return PodpisString;
        } //end function

         [CommandMethod("PrintPopFileCS", CommandFlags.Modal | CommandFlags.UsePickSet)]
        public void PrintPopFileCS()
        {
            //ФУНКЦИЯ ДЛЯ создания файла поперечников .рор. 
            BlockTable acBlkTbl;   //объявляем переменные для базы с примитивами чертежа 
            BlockTableRecord acBlkTblRec;
            //______создаем фильтр выбора линии сечения как 3-d полилинии чертежа____________________________________
            TypedValue[] tv = new TypedValue[1];
            tv.SetValue(new TypedValue((int)(DxfCode.Start), "POLYLINE"), 0);
            SelectionFilter filter = new SelectionFilter(tv);
            PromptSelectionOptions pso = new PromptSelectionOptions();
            pso.MessageForAdding = "\nВыберите 3-d полилинии поперечников";
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
                SelectionSet newSel = result.Value; // создаём переменную для записи в неё всех выбранных 3-д полилиний
                string sInfo;  // задаем переменную для информационной строки будущего файла
                string code = "0";
                string Opisanie = "";
                double RasstDoZero;
                double Htochki;
                double HforCode120 = -0.8;
                double HforCode130 = 3;
                double HforVX = 0.1;

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
                        acBlkTbl = (BlockTable)Trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                        acBlkTblRec = Trans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;  //открываем пространство модели для записи
                        try
                        {
                            Entity ent = Trans.GetObject(sObj.ObjectId, OpenMode.ForWrite) as Entity; //'приводим выбранный объект к типу Entity
                            Polyline3d MyPl3d = Trans.GetObject(ent.ObjectId, OpenMode.ForWrite) as Polyline3d; //приводим выбранный объект от типа entity к нужному мне типу Polyline3d

                            PromptResult PopNameResult = ed.GetString("\nВведите имя поперечника\n");//Ввод имени поперечника в дальнейшем планируется
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
                            PromptPointOptions ppo = new PromptPointOptions("\nУкажите нулевую точку на линии поперечника\n");
                            ppo.AllowNone = false;
                            PromptPointResult ZeroPointResult = ed.GetPoint(ppo);
                            Point3d ZeroPoint = ZeroPointResult.Value; //пользователь задает нулевую точку отсчета на поперечника (от нее "лево" и "право")
                            PolylineVertex3d Vert = new PolylineVertex3d();
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
                                    //Вычисляем дистанцию вдоль линии поперечника от вершины до нулевой точки.
                                    //Сейчас она вычисляется просто как длина отрезка.
                                    //!!!!!!!!!!!!Надо сделать как сумма длин сегментов ломаной линии!!!!!!!!!!______________________________________________________
                                    //RasstDoZero = Math.Round(Vychisli_S(Vert.Position, ZeroPoint), 2);
                                    RasstDoZero = Round(Vychisli_LomDlinu(MyPl3d, Vert.Position, ZeroPoint), 2);
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
                                            break;
                                        case "130":
                                            Htochki = HforCode130;
                                            break;
                                    }
                                    //____________________________________________________________________
                                    //запускаем процедуру печатания строки текстового файла поперечника
                                    sInfo = code + "      " + RasstDoZero.ToString() + "      " + Htochki.ToString() + "      " + Opisanie + "\n"; //записываем данные в строчку через разделитель пробелы
                                    writer.Write(sInfo); //записываем строку в файл
                                    //_________________________________________________________________________________________________________

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
        public double Vychisli_LomDlinu(Polyline3d myPolyline3d,Point3d myPoint,Point3d zeroPoint)
        {
            double LomDlina = 0;
            Point3dCollection myPoint3DCollection = new Point3dCollection();//создаем вспомогательную коллекцию, куда будем
                                                                            //кидать каждую точку анализируемой 3-д полилинии
            try
            {
                using (Transaction Trans = db.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId acObjIdVert in myPolyline3d)                               //перебираем каждую вершину 3-д полилинии, это делается как ObjectId в 3-д полилинии
                    {
                        PolylineVertex3d Vert;   // объявляем переменную для каждой 3-д вершины
                        Vert = Trans.GetObject(acObjIdVert, OpenMode.ForRead) as PolylineVertex3d; //считываем вершину по её ObjectId
                        Point3d anyPoint3D = new Point3d(Vert.Position.X, Vert.Position.Y, Vert.Position.Z);
                        myPoint3DCollection.Add(anyPoint3D);
                    }//кидаем следующую вершину
                }

                Point3dCollection lomPoint3DCollection = new Point3dCollection();

                int stInd = 0, finInd = 0;
                if (myPoint3DCollection.IndexOf(myPoint) < myPoint3DCollection.IndexOf(zeroPoint))
                {
                    stInd = myPoint3DCollection.IndexOf(myPoint);
                    finInd = myPoint3DCollection.IndexOf(zeroPoint);
                }
                if (myPoint3DCollection.IndexOf(zeroPoint) < myPoint3DCollection.IndexOf(myPoint))
                {
                    stInd = myPoint3DCollection.IndexOf(zeroPoint);
                    finInd = myPoint3DCollection.IndexOf(myPoint);
                }

                //функция вычисляет длину ломаной линии на плоскости, ограниченную заданными вершинами исходной 3-d полилинии
                int j;
                for (j = stInd; j <= finInd; j++)
                {
                    Point3d addPoint = new Point3d(myPoint3DCollection[j].X, myPoint3DCollection[j].Y, 0);
                    lomPoint3DCollection.Add(addPoint);
                }
                using (Polyline3d lomLine = new Polyline3d(Poly3dType.SimplePoly, lomPoint3DCollection, false))
                {
                    LomDlina = lomLine.Length;
                }
               
             
            }
           catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                ed.WriteMessage($"В процессе вычисления длины ломаной возникла ошибка: \n{ex}");

            }
            catch (System.Exception ex1)
            {
                ed.WriteMessage($"В процессе вычисления длины ломаной возникла ошибка: \n{ex1}");

            }
            return LomDlina;
        }//end function

        [CommandMethod("TestLomDlina", CommandFlags.Modal)]
        public void TestLomDlina()
        {
            PromptPointOptions ppo = new PromptPointOptions("\nВыберите точку");
            
            PromptPointResult ppo1 = ed.GetPoint(ppo);//выбираем первую точку
            
            if (ppo1.Status != PromptStatus.OK)
            {
                return;
            }
            PromptPointResult ppo2 = ed.GetPoint(ppo);//выбираем вторую точку
            if (ppo2.Status != PromptStatus.OK)
            {
                return;
            }
            Point3d myPoint = ppo1.Value;
            Point3d zeroPoint = ppo2.Value;

            PromptEntityOptions peoLine = new PromptEntityOptions("\nВыберите исходную 3-д полилинию\n");
            peoLine.SetRejectMessage("\nВыбрана не 3-д полилиния");
            peoLine.AddAllowedClass(typeof(Polyline3d), true);
            peoLine.AllowNone = false;
           
            PromptEntityResult myLineResult = ed.GetEntity(peoLine);//команда выбрать пользователю 3-д полилинию
            if (myLineResult.Status != PromptStatus.OK)
            {
                return;
            }
            using (Transaction Trans = db.TransactionManager.StartTransaction())
            {
                Polyline3d MyPl3d = Trans.GetObject(myLineResult.ObjectId, OpenMode.ForWrite) as Polyline3d;
                double lomanaya = Vychisli_LomDlinu(MyPl3d, myPoint, zeroPoint);
                ed.WriteMessage($"\nДлина ломаной по полилинии между точками составила {lomanaya}" );
                Trans.Commit();
            }
        }
        [CommandMethod("DlinaLomanoi", CommandFlags.Modal)]
        public void DlinaLomanoi()
        {
            ed.WriteMessage("\nКоманда вычисляет длину ломаного сегмента внутри выбранной \n(3д)полилинии, ограниченной двумя точками\n");
            PromptPointOptions ppo = new PromptPointOptions("\nВыберите точку");

            PromptPointResult ppo1 = ed.GetPoint(ppo);//выбираем первую точку

            if (ppo1.Status != PromptStatus.OK)
            {
                return;
            }
            PromptPointResult ppo2 = ed.GetPoint(ppo);//выбираем вторую точку
            if (ppo2.Status != PromptStatus.OK)
            {
                return;
            }
            Point3d myPoint = ppo1.Value;
            Point3d zeroPoint = ppo2.Value;
            PromptEntityOptions peoLine = new PromptEntityOptions("\nВыберите исходную полилинию (3-д полилинию)\n");
            peoLine.SetRejectMessage("\nВыбрана объект не того типа");
            peoLine.AddAllowedClass(typeof(Polyline3d), false);
            peoLine.AddAllowedClass(typeof(Polyline), false);
            peoLine.AllowNone = false;

            PromptEntityResult myLineResult = ed.GetEntity(peoLine);//команда выбрать пользователю полилинию
            if (myLineResult.Status != PromptStatus.OK)
            {
                return;
            }
 
            using (Transaction Trans = db.TransactionManager.StartTransaction())
            {
                Entity ent = (Entity)Trans.GetObject(myLineResult.ObjectId, OpenMode.ForRead);
                ed.WriteMessage("\nВыбран объект типа " + ent.GetType().ToString());
                switch (ent.GetType().ToString()) 

                {
                    case "Autodesk.AutoCAD.DatabaseServices.Polyline3d":
                   Polyline3d MyPl3d = Trans.GetObject(ent.ObjectId, OpenMode.ForWrite) as Polyline3d;
                double lomanaya = Vychisli_LomDlinu(MyPl3d, myPoint, zeroPoint);
                ed.WriteMessage($"\nДлина ломаной по полилинии между точками составила {lomanaya}");
                        break;
                    case "Autodesk.AutoCAD.DatabaseServices.Polyline":
                        Polyline MyPoly = Trans.GetObject(ent.ObjectId, OpenMode.ForWrite) as Polyline;
                        double lomanaya2 = MyCommonFunctions.Vychisli_LomDlinu_Poly(MyPoly, myPoint, zeroPoint);
                        ed.WriteMessage($"\nДлина ломаной по полилинии между точками составила {lomanaya2}");
                        break;
                }
               
                Trans.Commit();
            }
        }

        [CommandMethod("StartUFCsh", CommandFlags.Modal)]
        public void StartUFCsh()
        {
//_______процедура вставки пользовательского меню с панелями и кнопками
            RibbonButton ribbonButton1=new RibbonButton();
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
            ribbonButton3.Tag = ribbonButton2.Id;
            // привязываем к кнопке обработчик нажатия
            ribbonButton3.CommandHandler = new CommandHandler_ribbonButton3();

            // создаем контейнер для элементов
            Autodesk.Windows.RibbonPanelSource rbPanelSource = new Autodesk.Windows.RibbonPanelSource();
            rbPanelSource.Title = "Создание поперечников";
            // добавляем в контейнер элементы управления
            rbPanelSource.Items.Add(ribbonButton1);
            rbPanelSource.Items.Add(new RibbonSeparator());
            rbPanelSource.Items.Add(ribbonButton2);
            rbPanelSource.Items.Add(new RibbonSeparator());
            rbPanelSource.Items.Add(ribbonButton3);
            rbPanelSource.Items.Add(new RibbonSeparator());

            // создаем панель
            RibbonPanel rbPanel = new RibbonPanel();
            // добавляем на панель контейнер для элементов
            rbPanel.Source = rbPanelSource;

            // создаем вкладку
            RibbonTab rbTab = new RibbonTab();
            rbTab.Title = "Доп_функции";
            rbTab.Id = "AfoninRibbon";
            // добавляем на вкладку панель
            rbTab.Panels.Add(rbPanel);

            // получаем указатель на ленту AutoCAD
           RibbonControl rbCtrl = ComponentManager.Ribbon;
            // добавляем на ленту вкладку
            rbCtrl.Tabs.Add(rbTab);
            // делаем созданную вкладку активной ("выбранной")
            rbTab.IsActive = true;
        }
        [CommandMethod("Test_Risui_Poper", CommandFlags.Modal)]
        public static void Test_Risui_Poper()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor;
            Database db = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
            BlockTable acBlkTbl;   //объявляем переменные для базы с примитивами чертежа 
            BlockTableRecord acBlkTblRec;
            // //------Начинаем с создания нового пустого чертежа--------------
            // string strTemplatePath = "acadiso.dwt";
            // // DocumentCollection acDocMgr = Application.DocumentManager;
            // //Document acDoc = acDocMgr.Add(strTemplatePath);
            // //doc.Dispose();
            // Document popDoc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.Add(strTemplatePath);
            //// doc.Dispose();
            // Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument = popDoc;

            // Editor popEd = popDoc.Editor;

            // Database popDb= popDoc.Database;
            //_____________создан пустой файл - не получилось взаимодействовать с его Editor, оставил эту идею на потом______________________________
            //-----далее нужно считать текстовый файл ***.pop и вычертить линию поперечника в зависимости от масштабов-----
            // Запрашиваем имя pop-файла
            PromptOpenFileOptions pfo = new PromptOpenFileOptions("\nВыберите файл поперечника pop-файл: \n");
            pfo.Filter = "pop-файлы (*.pop)|*.pop";
            //pfo.DialogName = "Выбор файла поперечника";
            pfo.DialogCaption= "Выбор файла поперечника";
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
            Form1 form1= new Form1();
            //form1.Visible= true;
            //form1.groupBox2.Visible = true;
            if (Convert.ToDouble(form1.textBox1.Text)<1|| Convert.ToDouble(form1.textBox1.Text) > 1000)
            {
                ed.WriteMessage("\nВ поле масштаба введено неверное значение");
                return;
            }
            double horizScaleKoef = 1000 /Convert.ToDouble(form1.textBox1.Text);
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
                    Polyline profLine = new Polyline();//добавляем в базу полилинию, которая будет линией профиля
                    acBlkTblRec.AppendEntity(profLine);
                    Trans.AddNewlyCreatedDBObject(profLine, true);

                    string poperCaption = lines[0].Substring(5,lines[0].Length-5);
                    ed.WriteMessage($"\nСчитано значение заголовка поперечника: {poperCaption}\n");
                    //задаем стартовую точку для создания линии поперечника_____________________________________________
                    PromptPointOptions startPointOpt = new PromptPointOptions("\nУкажите стартовую точку 0 поперечника\n");
                    startPointOpt.AllowNone = false;
                    PromptPointResult startPointRes = ed.GetPoint(startPointOpt);
                    if (startPointRes.Status != PromptStatus.OK) return;
                    Point3d startPoint3D = startPointRes.Value;
                    //______________________Считываем из выбранного пользователем файла *.рор и раскидываем
                    //их по соответствующим массивам с текстами, строя основную линию профиля__________________________________
                    
                    bool changeZnak = true;
                    DBText[] dBTextsOpisanie = new DBText[] { };
                    DBText[] dBTextsOtmetki = new DBText[] { };
                    DBText[] dBTextsRasst = new DBText[] { };
                    DBText[] dBTextsCodes = new DBText[] { };

                   // double[,,] pointCodes = new double[,,] { };//может быть в перспективе буду использовать этот многомерный массив

                    double minY = 0;
                                        
                    for (int i = 1; i < lines.Length; i++) 
                    {   
                        if (lines[i][0]=='1'|| lines[i][0] == '0')//если строка начинается с символа 1 или 0 (то есть код есть), то разбиваем массив на подстроки
                        {
                        string[] znach;
                        znach = lines[i].Split(new char[] { ' ' },4,StringSplitOptions.RemoveEmptyEntries);//разбиваем строку на подстроки по разделитель "пробел",
                                                                                                           //максимум 4 подстроки (код,расстояние, высота, описание),
                                                                                                           //последовательные пробелы исключаются
                        
                        double Code = Convert.ToDouble(znach[0]);
                        double deltaX = Convert.ToDouble(znach[1]);//меняем знак расстояния в зависимости от "лево" или "право"
                            if (deltaX == 0)
                            {
                                changeZnak=false;
                            }
                            if (changeZnak == true)
                            {
                                deltaX = -deltaX;
                            }
                        double deltaY = Convert.ToDouble(znach[2]);
                            if (deltaY < minY) 
                            { 
                                minY = deltaY;
                            }
                         string opisanie = "";
                            if (znach.Length> 3)
                            {
                                opisanie = znach[3];
                            }
                            //ed.WriteMessage($"\nСчитано следующее содержимое поперечника:\n {Code}  {deltaX}  {deltaY} {opisanie}\n");

                            //в случае, если код точки 0, то в линию поперечника добавляем вершину_________________
                            
                            if (Code == 0)
                            {
                                
                                int indVert = 0;
                                Point2d poperVertex = new Point2d(startPoint3D.X+deltaX * horizScaleKoef, startPoint3D.Y+ deltaY * vertScaleKoef);
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
                                    newOpisanie.TextString = Convert.ToString(Abs(deltaX)) + " " + newOpisanie.TextString;
                                }
                                // newOpisanie.Height = 2;
                                newOpisanie.Height = Convert.ToDouble(form1.textBox3.Text);
                                newOpisanie.Rotation = 1.5708;
                                newOpisanie.Position = new Point3d(startPoint3D.X + deltaX * horizScaleKoef, startPoint3D.Y + deltaY * vertScaleKoef, 0);
                                dBTextsOpisanie.Append(newOpisanie);
                                
                            }
                            //_______________вставляем текст с отметкой точки, если больше 0__________________________
                            if (Code == 0 || Code == 1) 
                            {
                                DBText newOtmetka = new DBText();
                                acBlkTblRec.AppendEntity(newOtmetka);
                                Trans.AddNewlyCreatedDBObject(newOtmetka, true);
                                newOtmetka.TextString = Convert.ToString(Round(deltaY,2));
                                if (Code == 1) 
                                {
                                    newOtmetka.TextString = "СГР " + newOtmetka.TextString;
                                }
                                newOtmetka.Height = Convert.ToDouble(form1.textBox3.Text);
                                newOtmetka.Rotation = 1.5708;
                                newOtmetka.ColorIndex = 1;
                                newOtmetka.Position = new Point3d(startPoint3D.X + deltaX * horizScaleKoef, startPoint3D.Y + deltaY * vertScaleKoef-50, 0);
                                dBTextsOtmetki.Append(newOtmetka);

                            }
                            //_______________вставляем текст с расстоянием до 0-точки (далее мы их переделаем в расстояния между точками перелома)
                            // если код предусматривает другое - пропускаем__________________________
                            if (Code != 120 || Code != 121 || Code != 130 || Code != 131 || Code != 140 || Code != 141)
                            {
                                DBText newRasst = new DBText();
                                acBlkTblRec.AppendEntity(newRasst);
                                Trans.AddNewlyCreatedDBObject(newRasst, true);
                                newRasst.TextString = Convert.ToString(Round(deltaX, 2));

                                newRasst.Height = Convert.ToDouble(form1.textBox3.Text);
                                newRasst.Rotation = 1.5708;
                                newRasst.ColorIndex = 2;
                                newRasst.Position = new Point3d(startPoint3D.X + deltaX * horizScaleKoef, startPoint3D.Y + deltaY * vertScaleKoef - 60, 0);
                                dBTextsOtmetki.Append(newRasst);

                            }
                            
                            // сохраняем данные о кодах в массив текстов__________________________
                           
                                DBText newCode = new DBText();
                                acBlkTblRec.AppendEntity(newCode);
                                Trans.AddNewlyCreatedDBObject(newCode, true);
                                newCode.TextString = Convert.ToString(Code);
                                newCode.Height = Convert.ToDouble(form1.textBox3.Text);
                                newCode.Rotation = -1.5708;
                                newCode.ColorIndex = 3;
                                newCode.Position = new Point3d(startPoint3D.X + deltaX * horizScaleKoef, startPoint3D.Y + deltaY * vertScaleKoef - 70, 0);
                            dBTextsCodes.Append(newCode);

                            
                        }




                    }
                    MyCommonFunctions.ChertiPoper(profLine, minY, dBTextsOpisanie, dBTextsOtmetki, dBTextsRasst, dBTextsCodes, Trans, acBlkTblRec);

                        Trans.Commit();
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
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
    // обработчик нажатия кнопки
    public class CommandHandler_ribbonButton1 : System.Windows.Input.ICommand
    {
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object param)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("Sdelay_PoperCS" + " ", false, false, true);
        }
    }
    public class CommandHandler_ribbonButton2 : System.Windows.Input.ICommand
    {
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object param)
        {
            return true;
        }

        public void Execute(object parameter)
        {

            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("PrintPopFileCS" + " ", false, false, true);
        }
    }
    public class CommandHandler_ribbonButton3 : System.Windows.Input.ICommand
    {
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object param)
        {
            return true;
        }

        public void Execute(object parameter)
        {

            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("DlinaLomanoi" + " ", false, false, true);
        }
    }
}
       



    



           
       










