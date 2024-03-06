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
using System.Linq.Expressions;
using System.IO;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Drawing;
using System.Globalization;
using static System.Windows.Forms.LinkLabel;
using System.Reflection;
using System.Net;
using System.Runtime;

#if NCAD
using Teigha.DatabaseServices;
using Teigha.Runtime;
using Teigha.Geometry;
using HostMgd.ApplicationServices;
using HostMgd.EditorInput;
using HostMgd.Windows;
using Teigha.Colors;

using Platform = HostMgd;
using PlatformDb = Teigha;
#else
using Autodesk.Windows;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Internal;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices.Filters;
//using Autodesk.AutoCAD.Customization;
using Platform = Autodesk.AutoCAD;
    using PlatformDb = Autodesk.AutoCAD;
#endif

//using Autodesk.AutoCAD.Customization;
//using RibbonButton = Autodesk.AutoCAD.Customization.RibbonButton;
//using System.Activities.Expressions;

//[assembly: CommandClass(typeof(Useful_FunctionsCsh.AfoninCommands2))]


// This line is not mandatory, but improves loading performances

namespace Useful_FunctionsCsh
{
    public class Class1
    {
#if NCAD
        Document doc = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        Editor ed = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
     //   Editor ed= Teigha.Editor.CommandContext.Editor;
        Database db = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database;
#else
        Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
        Editor ed = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor;
        Database db = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
#endif
        //делаем функцию приведения всех характерных линий на отметку 0

        [CommandMethod("Sdelay_Zero", CommandFlags.Modal | CommandFlags.Session)]
        public void Sdelay_Zero()
        {
            //ObjectIdCollection objectIdAllEnts=new ObjectIdCollection();
            BlockTable acBlkTbl;   //объявляем переменные для базы с примитивами чертежа 
            BlockTableRecord acBlkTblRec;
            int counter = 0;
            TypedValue[] TvPoper = new TypedValue[2];
            TvPoper.SetValue(new TypedValue((int)(DxfCode.Start), "LWPOLYLINE,LINE,SPLINE"), 0);
            TvPoper.SetValue(new TypedValue((int)(DxfCode.LayerName), "1,2,3,4,5,6,7,8,10,12,13,14,20,23,24,25,28,29"), 1);
            SelectionFilter filterPoper = new SelectionFilter(TvPoper);
            PromptSelectionResult resultPoper = ed.SelectAll(filterPoper);
            SelectionSet poperSel = resultPoper.Value;
            if (poperSel != null)
            {

                foreach (SelectedObject sObj in poperSel) //перебираем каждый выбранный объект (3д-полилинию)
                {
                    using (Transaction Trans = db.TransactionManager.StartTransaction()) // начинаем транзакцию
                    {
                        DocumentLock docklock = doc.LockDocument();
                        acBlkTbl = (BlockTable)Trans.GetObject(db.BlockTableId, OpenMode.ForRead, false, true);      //открываем для чтения класс BlockTable
                        acBlkTblRec = Trans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false, true) as BlockTableRecord;
                        try
                        {
                            Entity anyEntity = (Entity)Trans.GetObject(sObj.ObjectId, OpenMode.ForWrite, false, true);
                            //if ((anyEntity.GeometricExtents.MinPoint.Z!=0)|| (anyEntity.GeometricExtents.MaxPoint.Z != 0))
                            //{
                            //    anyEntity.GeometricExtents.Set(new Point3d(anyEntity.GeometricExtents.MinPoint.X, anyEntity.GeometricExtents.MinPoint.Y, 0), new Point3d(anyEntity.GeometricExtents.MaxPoint.X, anyEntity.GeometricExtents.MaxPoint.Y, 0));
                            //    ++counter;
                            //}
                            switch (anyEntity.GetType().ToString())
                            {
#if NCAD
                                case "Teigha.DatabaseServices.Polyline":
#else
                                case "Autodesk.AutoCAD.DatabaseServices.Polyline":
#endif
                                    using (Polyline ProverkaPolyLine = Trans.GetObject(anyEntity.ObjectId, OpenMode.ForWrite) as Polyline)
                                    {
                                        if (ProverkaPolyLine.Elevation != 0)
                                        {
                                            ProverkaPolyLine.Elevation = 0;
                                            ++counter;
                                        }
                                    }
                                    break;
#if NCAD
                                case "Teigha.DatabaseServices.Line":
#else
                                case "Autodesk.AutoCAD.DatabaseServices.Line":
#endif
                                    using (Line ProverkaPolyLine = Trans.GetObject(anyEntity.ObjectId, OpenMode.ForWrite) as Line)
                                    {
                                        if ((ProverkaPolyLine.StartPoint.Z != 0) || (ProverkaPolyLine.EndPoint.Z != 0))
                                        {
                                            ProverkaPolyLine.StartPoint = new Point3d(ProverkaPolyLine.StartPoint.X, ProverkaPolyLine.StartPoint.Y, 0);
                                            ProverkaPolyLine.EndPoint = new Point3d(ProverkaPolyLine.EndPoint.X, ProverkaPolyLine.EndPoint.Y, 0);
                                            ++counter;
                                        }
                                    }
                                    break;
#if NCAD
                                case "Teigha.DatabaseServices.Spline":
#else
                                case "Autodesk.AutoCAD.DatabaseServices.Spline":
#endif
                                    using (Spline ProverkaPolyLine = Trans.GetObject(anyEntity.ObjectId, OpenMode.ForWrite) as Spline)
                                    {

                                        if ((ProverkaPolyLine.StartPoint.Z != 0) || (ProverkaPolyLine.EndPoint.Z != 0))
                                        {
                                            ProverkaPolyLine.StartPoint = new Point3d(ProverkaPolyLine.StartPoint.X, ProverkaPolyLine.StartPoint.Y, 0);//!!!!
                                            ProverkaPolyLine.EndPoint = new Point3d(ProverkaPolyLine.EndPoint.X, ProverkaPolyLine.EndPoint.Y, 0);
                                            ++counter;
                                        }
                                    }
                                    break;

                            }
                            docklock.Dispose();
                            Trans.Commit();
                        }
#if NCAD
                        catch (Teigha.Runtime.Exception ex)
#else
                        catch (Autodesk.AutoCAD.Runtime.Exception ex)
#endif
                        //в случае обнаружения ошибки пишем её описание и прерываем транзакцию

                        {
                            ed.WriteMessage("В процессе приведения линий на отметку 0 возникла ошибка:\n " + ex.Message);
                            Trans.Abort();
                        }
                        catch (System.Exception ex1)
                        //в случае обнаружения ошибки пишем её описание и прерываем транзакцию

                        {
                            ed.WriteMessage("В процессе приведения линий на отметку 0 возникла ошибка:\n " + ex1.Message);
                            Trans.Abort();
                        }
                    }
                }
            }
            ed.WriteMessage($"Внесены изменения в геометрию {counter} объектов\n");



        }
        public Point3d Find_FirstPoint(Entity myEntity, Transaction transaction)
        {
            //функция находит первую точку у полилинии, отрезка, сплайна
            Point3d first_point3D = new Point3d();
            switch (myEntity.GetType().ToString())
            {
#if NCAD
                case "Teigha.DatabaseServices.Polyline":
#else
                case "Autodesk.AutoCAD.DatabaseServices.Polyline":
#endif
                    Polyline polyline = (Polyline)transaction.GetObject(myEntity.ObjectId, OpenMode.ForRead);
                    first_point3D = new Point3d(polyline.GetPoint2dAt(0).X, polyline.GetPoint2dAt(0).Y, 0);
                    break;
#if NCAD
case "Teigha.DatabaseServices.Line":
#else
                case "Autodesk.AutoCAD.DatabaseServices.Line":
#endif
                    Line line = (Line)transaction.GetObject(myEntity.ObjectId, OpenMode.ForRead);
                    first_point3D = line.StartPoint;
                    break;
#if NCAD
                case "Teigha.DatabaseServices.Spline":
#else
                case "Autodesk.AutoCAD.DatabaseServices.Spline":
#endif
                    Spline spline = (Spline)transaction.GetObject(myEntity.ObjectId, OpenMode.ForRead);
                    first_point3D = spline.StartPoint;
                    break;

            }

            return first_point3D;
        }
        public ObjectId Find_CurDescription(Point3d myPoint, Transaction Trans)
        {
            ObjectId id = new ObjectId();
            //функция находит текст с текущим заданным описанием
            TypedValue[] TvOpisanie = new TypedValue[2];
            TvOpisanie.SetValue(new TypedValue((int)(DxfCode.Start), "TEXT"), 0);
            TvOpisanie.SetValue(new TypedValue((int)(DxfCode.LayerName), "подписи поперечников"), 1);
            SelectionFilter filterOpisanie = new SelectionFilter(TvOpisanie);
            PromptSelectionResult resultOpisanie = ed.SelectAll(filterOpisanie);
            SelectionSet poperOpisanie = resultOpisanie.Value;
            DBText TryOpisanie = new DBText();

            foreach (SelectedObject ObjOpisanie in poperOpisanie)
            {
                TryOpisanie = Trans.GetObject(ObjOpisanie.ObjectId, OpenMode.ForRead, false, false) as DBText;
                if (TryOpisanie.Position.X == myPoint.X & TryOpisanie.Position.Y == myPoint.Y)
                {
                    id = TryOpisanie.ObjectId;
                    break;
                } //End If


            } //Next
            return id;
        }


        [CommandMethod("Make_Description", CommandFlags.Modal | CommandFlags.UsePickSet | CommandFlags.Session)]
        public void Make_Description()
        {
            BlockTable acBlkTbl;   //объявляем переменные для базы с примитивами чертежа 
            BlockTableRecord acBlkTblRec;
            using (Transaction Trans = db.TransactionManager.StartTransaction()) // начинаем транзакцию
            {
                DocumentLock docklock = doc.LockDocument(DocumentLockMode.AutoWrite, null, null, true);
                acBlkTbl = (BlockTable)Trans.GetObject(db.BlockTableId, OpenMode.ForRead, false, true);      //открываем для чтения класс BlockTable
                acBlkTblRec = Trans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false, true) as BlockTableRecord;
                TypedValue[] TvPoper = new TypedValue[2];
                TvPoper.SetValue(new TypedValue((int)(DxfCode.Start), "LWPOLYLINE,LINE,SPLINE"), 0);
                TvPoper.SetValue(new TypedValue((int)(DxfCode.LayerName), "1,2,3,4,5,6,7,8,10,12,13,14,20,23,24,28,29"), 1);
                SelectionFilter filterPoper = new SelectionFilter(TvPoper);
                PromptSelectionOptions promptSelectionOptions = new PromptSelectionOptions();
                promptSelectionOptions.MessageForAdding = "\nВыберите характерную линию\n";
                promptSelectionOptions.SingleOnly = true;
                PromptSelectionResult resultPoper = ed.GetSelection(promptSelectionOptions, filterPoper);
                if (resultPoper.Status == PromptStatus.OK)
                {
                    //_____________________с помощью транзакции добавляем в базу чертежа слой "подписи поперечников" (а если он там уже есть - то не добавляем)______________________________
                    LayerTable acLyrTbl = (LayerTable)Trans.GetObject(db.LayerTableId, OpenMode.ForRead);
                    String sLayerName = "подписи поперечников";

                    if (acLyrTbl.Has(sLayerName) == false)
                    {
                        LayerTableRecord acLyrTblRec = new LayerTableRecord();
                        // Устанавливаем слою нужный мне цвет по индексу 190 и ранее заданное имя
#if NCAD
                        acLyrTblRec.Color =Teigha.Colors.Color.FromColorIndex(ColorMethod.ByAci, 190);

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
                    SelectionSet poperSel = resultPoper.Value;
                    ObjectId myObjectId = poperSel.GetObjectIds()[0];
                    Entity myEnt = (Entity)Trans.GetObject(myObjectId, OpenMode.ForWrite);
                    String myDescription = SdelayOpisanieTochkiPopera_3(myEnt, Trans);
                    //  ed.WriteMessage($"\nНашел описание линии:{myDescription}\n");
                    DBText opisText = new DBText();
                    Point3d opisTextInsPt = Find_FirstPoint(myEnt, Trans);
                    // ed.WriteMessage($"Координаты первой точки выбранной линии: {opisTextInsPt}\n");
                    bool ExistDescripInPoint = true;
                    ExistDescripInPoint = ProveritNalichiePodpisiInTochka(opisTextInsPt);
                    //  ed.WriteMessage($"Существует в начальной точке текущее описание: {ExistDescripInPoint}\n");
                    if (ExistDescripInPoint)
                    {
                        docklock.Dispose();
                        ObjectId idOpisText = Find_CurDescription(opisTextInsPt, Trans);
                        DBText old_opisText = (DBText)Trans.GetObject(idOpisText, OpenMode.ForWrite);
                        ed.WriteMessage($"Найдено старое описание {old_opisText.TextString}\n");
                        //opisText.TextString= myDescription;
                        old_opisText.Erase(true);
                        old_opisText.Dispose();


                    }

                    //------------------------------
                    opisText.TextString = myDescription;
                    opisText.Position = opisTextInsPt;
                    opisText.Layer = "подписи поперечников";
                    opisText.Justify = AttachmentPoint.BaseLeft;
                    opisText.Rotation = 0;
                    opisText.Height = 0.5;
                    opisText.ColorIndex = 190;
                    opisText.LineWeight = (LineWeight)0.3;
                    acBlkTblRec.AppendEntity(opisText);
                    Trans.AddNewlyCreatedDBObject(opisText, true);


                    //далее создаю диалог с выскакивающим окном для ввода желаемого текста описания.
                    //myDescription используется как текущий вариант по-умолчанию
                    //Form2 form2 = new Form2();

                    //Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(form2);
                    //form2.textBox1.Text = myDescription;
                    //DialogResult dialogOpis =form2.DialogResult;
                    //if (dialogOpis == DialogResult.Yes)
                    //{
                    //    myDescription = dialogOpis.ToString();
                    //    // myDescription = form2.returnString;
                    //}
                    //___________________________________________________________________
                    PromptStringOptions optionsOpis = new PromptStringOptions("Введите желаемое описание точки\n");
                    optionsOpis.DefaultValue = myDescription;
                    optionsOpis.UseDefaultValue = true;
                    optionsOpis.AllowSpaces = true;

                    PromptResult resultOpis = ed.GetString(optionsOpis);
                    if (resultOpis.Status != PromptStatus.OK)
                    {
                        ed.WriteMessage("Прервана работа команды\n");
                        return;
                    }
                    myDescription = resultOpis.StringResult;


                    opisText.TextString = myDescription;


                    //ed.WriteMessage($"\nБудем создавать описание {myDescription}\n");
                }
                else
                {
                    ed.WriteMessage("\nНе выбраны линии для создания подписи\n");
                }


                Trans.Commit();
                //docklock.Dispose();
            }



        }
        public int find_addvertex_index(Polyline myPolyline, Point3d addedPoint)
        {
            //функция возвращает индекс вершины,
            //перед которой можно вставить в полилинию проверяемую точку, чтобы она не "перекрутилась" (по порядку)
            int numer = 0;
            for (int i = 0; i < myPolyline.NumberOfVertices - 1; ++i)
            {

                double S1 = Vychisli_S(myPolyline.GetPoint3dAt(i), myPolyline.GetPoint3dAt(i + 1));
                double S2 = Vychisli_S(myPolyline.GetPoint3dAt(i), addedPoint);
                if (S2 <= S1)
                {
                    numer = i + 1; break;
                }
            }
            return numer;
        }
        public bool is_point_in_Current_Vertex(Point3d point, Polyline polyline)
        {
            bool point_vertex = false;
            for (int i = 0; i < polyline.NumberOfVertices; ++i)
            {
                Point3d vertPoint = polyline.GetPoint3dAt(i);
                if (vertPoint == point)
                {
                    point_vertex = true;
                    break;
                }
            }
            return point_vertex;
        }

        [CommandMethod("group_intersection", CommandFlags.UsePickSet)]
        public void group_intersection()
        {
            //команда должна предлагать выбрать пользователю 1 секущую полилинию и на основании автоматического фильтра пересекаемых полилиний,
            // должна создавать дополнительную вершину с правильным порядковым индексом у каждой пересекаемой линии в месте пересечения (+ опционально 3-д точку)

            BlockTable acBlkTbl;   //объявляем переменные для базы с примитивами чертежа 
            BlockTableRecord acBlkTblRec;
            //______создаем фильтр выбора секущей линии ____________________________________
            TypedValue[] tv = new TypedValue[1];
            tv.SetValue(new TypedValue((int)(DxfCode.Start), "LWPOLYLINE"), 0); //проблема здесь
            SelectionFilter filter = new SelectionFilter(tv);
            PromptSelectionOptions pso = new PromptSelectionOptions();
            pso.MessageForAdding = "\nВыберите секущую полилинию\n";
            pso.SingleOnly = true;
            pso.SinglePickInSpace = false;
            //_______________________________________________________________________________________________________________
            PromptSelectionResult result = ed.GetSelection(pso, filter); //команда пользователю выбрать на экране секущую полилинию, выбор идёт в соответствии с фильтром и опциями
            if (result.Status == PromptStatus.OK) // в случае, если пользователь выбрал линию, то идем дальше
            {
                SelectionSet cuttingSel = result.Value; // создаём переменную для записи в неё выбранного набора

                TypedValue[] TvPoper = new TypedValue[2];
                TvPoper.SetValue(new TypedValue((int)(DxfCode.Start), "LWPOLYLINE"), 0);
                TvPoper.SetValue(new TypedValue((int)(DxfCode.LayerName), "20,23,24,25,28,29,_survey жд пути,_survey балласт,Здания и строения,Инженерные сооружения,Путевое хозяйство,Автодорожное хозяйство,Объекты электропередачи,Гидрография,Откосы,Рельеф,Растительность,Ограждения,06_Инженерно-технические сооружения,03_Здания и строения,10_Границы покрытий и угодий,05_Элементы зданий,09_Путевое хозяйство,07_Объекты электропередачи,11_Гидрография,12_Рельеф,13_Растительность,14_Ограждения"), 1);
                SelectionFilter filterPoper = new SelectionFilter(TvPoper);
                PromptSelectionResult resultPoper = ed.SelectAll(filterPoper);
                SelectionSet poperSel = resultPoper.Value;
                if (poperSel != null)
                {
                    using (Transaction Trans = db.TransactionManager.StartTransaction()) // начинаем транзакцию
                    {
                        

                        acBlkTbl = (BlockTable)Trans.GetObject(db.BlockTableId, OpenMode.ForRead, false, true);      //открываем для чтения класс BlockTable
                        acBlkTblRec = Trans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false, true) as BlockTableRecord;

                        Polyline polyline_Cutting = (Polyline)Trans.GetObject(cuttingSel.GetObjectIds()[0], OpenMode.ForRead, false, true);
                        if (polyline_Cutting != null)
                        {
                            ObjectIdCollection int_pol_id= new ObjectIdCollection();
                            //сначала удаляем из набора те полилинии, которые не пересекаются с секущей
                            foreach (SelectedObject sObj in poperSel) //перебираем каждый выбранный объект (пересекаемую линию)
                            {
                                Polyline polyline_Intersected = (Polyline)Trans.GetObject(sObj.ObjectId, OpenMode.ForRead, false, true);
                                Point3dCollection intersect_point3dCol = new Point3dCollection();
                                polyline_Cutting.IntersectWith(polyline_Intersected, Intersect.OnBothOperands, intersect_point3dCol, IntPtr.Zero, IntPtr.Zero);
                                if (intersect_point3dCol.Count > 0)
                                {
                                    int_pol_id.Add(sObj.ObjectId);
                                }
                            }
                            if (int_pol_id.Count > 0)
                            {
                                foreach (ObjectId sObjId in int_pol_id) //перебираем каждый выбранный объект (пересекаемую линию)
                                {
                                    Polyline polyline_Intersected = (Polyline)Trans.GetObject(sObjId, OpenMode.ForRead, false, true);
                                    if ((polyline_Intersected != null))
                                    {
                                        if ((polyline_Intersected.Layer == "25"|| polyline_Intersected.Layer == "Откосы"|| polyline_Intersected.Layer == "12_Рельеф") && (polyline_Intersected.NumberOfVertices < 3)) continue;
                                        added_Vertex_Polyline(polyline_Cutting, polyline_Intersected);
                                        //  ed.WriteMessage($"Функция \"added_Vertex_Polyline\" вроде отработала успешно\n");

                                    }
                                    else
                                    {
                                        ed.WriteMessage("Не найдено ObjectId пересекаемой полилинии\n");
                                        Trans.Abort();
                                    }

                                }//берем следующий объект из пересекаемого набора
                            }
                            

                        }

                        Trans.Commit();
                    }
                }else ed.WriteMessage("В стандартный набор не попало ни одной пересекаемой полилинии\n");

            }
            else ed.WriteMessage("Не выбрана секущая полилиния\n");

        }
        public void added_Vertex_Polyline(Polyline polyline_Cutting, Polyline polyline_Intersected)
        {
            if (polyline_Cutting.StartPoint.Z != polyline_Intersected.StartPoint.Z)
            {
                ed.WriteMessage($"Секущая полилиния с пересекаемой полилинией с ObjectId {polyline_Intersected.ObjectId} находятся на разных уровнях\n");
            }
            //  Polyline modified_Polyline=new Polyline();
            int myIndex = 0;
            bool need_DBPoint = false;
            Form1 form1 = new Form1();
            
            //if (form1.checkBox2.Checked && form1.checkBox2.CheckState==CheckState.Checked)
                if (form1.checkBox2.Checked)
                {
                      need_DBPoint=true;
                }
            Point3dCollection intersect_point3dCol = new Point3dCollection();
            polyline_Cutting.IntersectWith(polyline_Intersected, Intersect.OnBothOperands, intersect_point3dCol, IntPtr.Zero, IntPtr.Zero);
            if (intersect_point3dCol.Count > 0) // если пересечения нашлись - идем дальше, иначе сообщаем, что не нашли пересечений
            {
                // ed.WriteMessage($"С линией с ObjectId: {polyline_Intersected.ObjectId} найдено {intersect_point3dCol.Count} точек пересечения:\n");
                //пока создадим вспомогательное сообщение с количеством и координатами найденных точек
                foreach (Point3d int_point in intersect_point3dCol)
                {
                    // ed.WriteMessage($"\t{int_point}\n");
                    // сначала проверяем, не совпала ли точка пересечения с существующей вершиной полилинии
                    // - в этом случае ничего добавлять не нужно
                    bool isVertex = is_point_in_Current_Vertex(int_point, polyline_Intersected);
                    if (isVertex == false)
                    {

                        //в отдельной функции находим индекс, после которого надо вставить в пересекаемую линию вершину
                        myIndex = find_addvertex_index(polyline_Intersected, int_point);
                        polyline_Intersected.UpgradeOpen();
                        polyline_Intersected.AddVertexAt(myIndex, new Point2d(int_point.X, int_point.Y), 0, 0, 0);
                        // ed.WriteMessage($"В пересекаемую полилинию добавлена вершина перед вершиной {myIndex}\n");

                    }
                    else
                    {
                        //если точка пересечения совпала с текущей вершиной, то ничего не делаем, переходим к следующей точке в коллекции
                        ed.WriteMessage($"Точка пересечения {int_point} совпала с существующей вершиной пересекаемой полилинии\n");
                       // continue;
                    }
                    //----------------------------начальная задумка сделана. Далее (опционально, если это указано в управляющей форме,
                    //---------------------вставляем вершину и в секущую полилинию, и создаем 3-Д точки на месте пересечений с отметками,
                    //-------------------------------------------интерполированными по пересекаемым линиям-----------------------------------
                   bool isVertex_Cutting = is_point_in_Current_Vertex(int_point, polyline_Cutting);
                    if (isVertex_Cutting == false)
                    {
                       int cutting_Index = find_addvertex_index(polyline_Cutting, int_point);
                        polyline_Cutting.UpgradeOpen();
                        polyline_Cutting.AddVertexAt(cutting_Index, new Point2d(int_point.X, int_point.Y), 0, 0, 0);
                        //  ed.WriteMessage($"В секущую полилинию добавлена вершина перед вершиной {myIndex}\n");

                    }
                    else
                    {
                        ed.WriteMessage($"Точка пересечения {int_point} совпала с существующей вершиной секущей полилинии\n");
                       // continue;
                    }
                    //опционально создаем 3-д точки
                    //--------далее запускаем процедуру создания точки в полученном пересечении----------------------
                    //--для создания точки с отметкой у меня есть ее Х и Y (int_point),
                    //для вычисления отметки надо найти на нужной структурной линии (polyline_Intersected) ее такую вершину,
                    //у которой есть существующая DBPoint с отметкой (по линии дальше и по линии ближе
                    //к началу от найденной точки пересечения)--------------------------------------------------------
                    if (need_DBPoint)
                    {
                        // if (isVertex) myIndex= polyline_Intersected.GetPoint3dAt(int_point)
                        if (isVertex) myIndex = find_addvertex_index(polyline_Intersected, int_point);
                        createPointOnPolyline(int_point, polyline_Intersected, myIndex); // решить проблему получения индекса точки пересечения в существующей вершине
                    }
                    else ed.WriteMessage("Процедура создания точек пропущена из-за настроек\n");
                    
                }

            }
            else
            {
                ed.WriteMessage($"Не нашлось пересечений с линией с ObjectId: {polyline_Intersected.ObjectId}\n");
                // return null;
            }
            //  return modified_Polyline;
        }
        public void createPointOnPolyline(Point3d int_point, Polyline polyline_Intersected, int myIndex)
        {

            Point3dCollection all_points = new Point3dCollection();
            BlockTable acBlkTbl;   //объявляем переменные для базы с примитивами чертежа 
            BlockTableRecord acBlkTblRec;
            using (Transaction Trans = db.TransactionManager.StartTransaction()) // начинаем транзакцию
            {
                acBlkTbl = (BlockTable)Trans.GetObject(db.BlockTableId, OpenMode.ForRead, false, true);      //открываем для чтения класс BlockTable
                acBlkTblRec = Trans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false, true) as BlockTableRecord;//теперь мы получили доступ к пространству модели
                TypedValue[] TvPoints = new TypedValue[1];
                TvPoints.SetValue(new TypedValue((int)(DxfCode.Start), "POINT"), 0);
                SelectionFilter filterPoints = new SelectionFilter(TvPoints);
                PromptSelectionResult resultPoints = ed.SelectAll(filterPoints);
              //  PromptSelectionResult resultPoints = ed.SelectCrossingWindow(new Point3d(int_point.X + 100, int_point.Y - 100, 0), new Point3d(int_point.X - 100, int_point.Y + 100, 0), filterPoints);
                if (resultPoints.Status == PromptStatus.OK)
                {
                    foreach (SelectedObject Sobj in resultPoints.Value)// закидываем все точки чертежа в общую коллекцию
                    {
                        DBPoint selected_DBPoint = (DBPoint)Trans.GetObject(Sobj.ObjectId, OpenMode.ForRead,false);
                        if ((selected_DBPoint.Position == int_point)&&(selected_DBPoint.IsErased==false))
                        {
                            ed.WriteMessage($"У полилинии на слое { polyline_Intersected.Layer} определяемая точка совпала с текущей точкой в чертеже\n");
                        }
                        else
                        {
                            all_points.Add(selected_DBPoint.Position);
                        }
                    }

                } else
                {
                    ed.WriteMessage("В текущем чертеже не найдено 3-d-точек\n");
                    Trans.Abort();
                }
                //после того, как заполнена коллекция (если в ней больше 1 объекта, идем по полилинии и ищем совпадение ее вершины с какой-нибудь точкой и если находим - берем ее отметку
              //  ed.WriteMessage($"В анализ берется {all_points.Count} точек\n");
                if (all_points.Count > 1)
                {
                    Point3d point3D_after = new Point3d();
                    Point3d point3D_before = new Point3d();
                    int count_made = 0;
                    for (int i = myIndex + 1; i < polyline_Intersected.NumberOfVertices; ++i)
                    {
                        point3D_after = find_Z(polyline_Intersected.GetPoint3dAt(i), all_points);
                        if (point3D_after.Z > double.MinValue) break;
                    }
                        if (point3D_after.Z > double.MinValue)
                        {
                            for (int j = myIndex - 1; j >= 0; --j)
                            {
                                point3D_before = find_Z(polyline_Intersected.GetPoint3dAt(j), all_points);
                            if (point3D_before.Z > double.MinValue) break;
                            }
                                if (point3D_before.Z > double.MinValue)
                                {
                                    //после того, как нам стали известны 3-d точки на линии до и после пересечения, интерполяцией нужно найти отметку точки пересечения
                                    double myZ = Vychisli_Z_on_Polyline(point3D_before, int_point, point3D_after, polyline_Intersected);
                                    Point3d point3Dinterpol = new Point3d(int_point.X, int_point.Y, myZ);
                                    DBPoint int_DPpoint = new DBPoint(point3Dinterpol);
                                    acBlkTblRec.AppendEntity(int_DPpoint);
                                    Trans.AddNewlyCreatedDBObject(int_DPpoint, true);
                                    Trans.Commit();
                                    ++count_made;
                                   // return;
                                   // break;
                                }
                                //else
                                //{
                                //    ed.WriteMessage($"У полилинии на слое {polyline_Intersected.Layer} не найдено подходящих точек до точки пересечения\n");
                                //  //  Trans.Abort();
                                //}

                            
                           // break;  

                        }
                        //else
                        //{
                        //    ed.WriteMessage($"У полилинии на слое {polyline_Intersected.Layer} не найдено подходящих точек после точки пересечения\n");
                        //   // Trans.Abort();
                        //}
                    
                    if (count_made == 0) ed.WriteMessage($"Для полилинии в слое \"{polyline_Intersected.Layer}\" не создано ни одной точки в месте пересечения\n");
                }
                else
                {
                    ed.WriteMessage("В текущем чертеже слишком мало 3-d-точек\n");
                    Trans.Abort();
                }
            }
        }
        public double Vychisli_Z_on_Polyline(Point3d point3D_before, Point3d int_point, Point3d point3D_after, Polyline polyline_Intersected)
        {
            double S1 = MyCommonFunctions.Vychisli_LomDlinu_Poly(polyline_Intersected,new Point3d(point3D_before.X, point3D_before.Y,0),new Point3d(int_point.X, int_point.Y,0));
           // ed.WriteMessage($"Для анализа точки пересечения найдена предыдущая точка {point3D_before}\nПолучено расстояние по кривой от предыдущей точки {S1}\n");
            double S2 = MyCommonFunctions.Vychisli_LomDlinu_Poly(polyline_Intersected, new Point3d(point3D_after.X, point3D_after.Y, 0), new Point3d(int_point.X, int_point.Y, 0));
          //  ed.WriteMessage($"Для анализа точки пересечения найдена последующая точка {point3D_after}\nПолучено расстояние по кривой до последующей точки {S2}\n");
            double Hfull = point3D_after.Z - point3D_before.Z;
            double Spol = S1 + S2;
            double TanA = Hfull / Spol;
            double H1 = S1 * TanA;
          //  ed.WriteMessage($"Вычислена высота точки пересечения {point3D_before.Z + H1}\n");
            return point3D_before.Z+H1;
        }
        public Point3d find_Z(Point3d anyPoint3D,Point3dCollection sourceCol)
        {
            double myZ = double.MinValue;
           // Point3d point3D_onVert = new Point3d(0,0,myZ);
            foreach (Point3d col in sourceCol)
            {
                if (Round(anyPoint3D.X,3)==Round(col.X,3)&&Round(anyPoint3D.Y,3)==Round(col.Y,3))
                {
                    myZ = col.Z;
                    break;
                }
            }
            //ed.WriteMessage($"Получена точка с отметкой {myZ}\n");
            return new Point3d(anyPoint3D.X, anyPoint3D.Y, myZ);
        }
        public Point3d get_middle_point3D(Point3d point3D_1, Point3d point3D_2)
        {
            double dX=point3D_2.X- point3D_1.X;
            double dY = point3D_2.Y - point3D_1.Y;
            double dZ = point3D_2.Z - point3D_1.Z;
            Point3d middle_point3D = new Point3d(point3D_1.X + 0.5 * dX, point3D_1.Y + 0.5 * dY, point3D_1.Z + 0.5 * dZ);
            return middle_point3D;
        }
        public double get_angle_90(Point3d point3D_1, Point3d point3D_2)
        {
            double angle_90 = 0;
            using (Line myLine=new Line())
            {
                myLine.StartPoint=point3D_1;
                myLine.EndPoint=point3D_2;
                angle_90=myLine.Angle+1.5708;
            }
            return angle_90;
        }
        public Point3dCollection delete_dubles(Point3dCollection point3DCollection_with_doubles)
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
        public void SdelayMPtexts(Point3dCollection SortPointsGRcol, string textMP_style, double textMP_height, string textMP_layer)
        {

            //метод создает тексты междопутий. Нужный слой уже создан в базе (перед выызовом процедуры).
            BlockTable acBlkTbl;   //объявляем переменные для базы с примитивами чертежа 
            BlockTableRecord acBlkTblRec;
          //  DocumentLock docklock = doc.LockDocument(DocumentLockMode.ExclusiveWrite,null,null,true);
            using (Transaction Trans = db.TransactionManager.StartTransaction()) // начинаем транзакцию
            {
                try //начинаем обработку с блоком конструкцией улавливания ошибок
                {
                    //Если есть указанный стиль в базе, то делать им.
                    acBlkTbl = (BlockTable)Trans.GetObject(db.BlockTableId, OpenMode.ForRead, false, true);      //открываем для чтения класс BlockTable
                    acBlkTblRec = Trans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false, true) as BlockTableRecord;//теперь мы получили доступ к пространству модели
                    //сначала получаем примерное положение текста посередине между точками коллекции
                    for (int i = 0;i< SortPointsGRcol.Count-1; ++i)
                    {
                        Point3d point3D_1 = SortPointsGRcol[i];
                        Point3d point3D_2 = SortPointsGRcol[i+1];
                        Point3d point_ins_MPtext = get_middle_point3D(point3D_1, point3D_2);//находим среднюю точку с помощью вспомогательной функции
                    ed.WriteMessage("Функция \"get_middle_point3D\" для середины междопутья отработала  успешно\n");  
                    //для создания красивого текста междопутья осталось узнать угол поворота текста и его содержимое (расстояние между точками)
                        double text_MP_angle = get_angle_90(point3D_1, point3D_2);
                    ed.WriteMessage("Функция \"get_angle_90\" отработала успешно\n");
                    double text_MP_rasst = Vychisli_S(point3D_1, point3D_2);
                        string text_MP_string= text_MP_rasst.ToString("0.00", CultureInfo.InvariantCulture);//задаем содержимое правильного формата
                                                                                                            //все исходные данные известны, создаем текст в примерном положении и добавляем его в базу чертежа
                        TextStyleTable textStyleTableDoc = (TextStyleTable)Trans.GetObject(db.TextStyleTableId, OpenMode.ForRead, false, true);//открываем для чтения таблицу стилей текста
                        String sStyleName = textMP_style;

                        if (textStyleTableDoc.Has(sStyleName))
                        {
                            db.Textstyle = textStyleTableDoc[sStyleName];

                        }

                    DBText text_MP=new DBText();
                    acBlkTblRec.AppendEntity(text_MP);
                    Trans.AddNewlyCreatedDBObject(text_MP, true);
                    ed.WriteMessage($"Текст без содержимого создан  успешно\n");
                    //устанавливаем текущим стиль текста, переданный в функции, если он есть в базе чертежа
                   
                    text_MP.TextString= text_MP_string;
                        text_MP.Rotation = text_MP_angle;
                        text_MP.Position= point_ins_MPtext;
                        text_MP.Layer = textMP_layer;
                        text_MP.Height = textMP_height;
                        text_MP.ColorIndex = 256;
                        text_MP.LineWeight = LineWeight.ByLayer;

                        //теперь перемещаем созданный текст, чтобы он был левее линии поперечника, и его середина была посередине линии
                        Vector3d myVector = text_MP.GeometricExtents.MinPoint.GetVectorTo(text_MP.Position);
                        text_MP.TransformBy(Matrix3d.Displacement(myVector));
                        Point3d middle_box = get_middle_point3D(text_MP.GeometricExtents.MaxPoint, text_MP.Position);
                        myVector = middle_box.GetVectorTo(text_MP.Position);
                        text_MP.TransformBy(Matrix3d.Displacement(myVector));

                        if (text_MP.TextStyleName == "ATP")
                        {
                            text_MP.Oblique = 0.2618;
                            text_MP.WidthFactor = 0.8;
                        }
                    }

                    Trans.Commit();

                }
#if NCAD
                catch (Teigha.Runtime.Exception ex)
#else
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
#endif
                {
                    ed.WriteMessage($"В процессе работы команды \"SdelayMPtexts\" обнаружена ошибка: {ex.Message}\n");
                    Trans.Abort();
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"В процессе работы команды \"SdelayMPtexts\" обнаружена ошибка: {ex.Message}\n");
                    Trans.Abort();
                }
            }

        }
        [CommandMethod("Test_Intersection", CommandFlags.UsePickSet)]
        public void Test_Intersection()
        {
            //команда должна предлагать выбрать пользователю 1 секущую полилинию и некоторое количество (>0) пересекаемых полилиний,
            //и должна создавать дополнительную вершину с правильным порядковым индексом у каждой пересекаемой линии в месте пересечения

            BlockTable acBlkTbl;   //объявляем переменные для базы с примитивами чертежа 
            BlockTableRecord acBlkTblRec;
            //______создаем фильтр выбора секущей линии ____________________________________
            TypedValue[] tv = new TypedValue[1];
            tv.SetValue(new TypedValue((int)(DxfCode.Start), "LWPOLYLINE"), 0); //проблема здесь
            SelectionFilter filter = new SelectionFilter(tv);
            PromptSelectionOptions pso = new PromptSelectionOptions();
            pso.MessageForAdding = "\nВыберите секущую полилинию\n";
            pso.SingleOnly = true;
            pso.SinglePickInSpace = false;
            //_______________________________________________________________________________________________________________
            PromptSelectionResult result = ed.GetSelection(pso, filter); //команда пользователю выбрать на экране секущую полилинию, выбор идёт в соответствии с фильтром и опциями
            if (result.Status == PromptStatus.OK) // в случае, если пользователь выбрал линию, то идем дальше
            {
                SelectionSet cuttingSel = result.Value; // создаём переменную для записи в неё выбранного набора
                // далее предлагаем выбрать несколько пересекаемых полилиний
                PromptSelectionOptions pso_1 = new PromptSelectionOptions();
                pso_1.MessageForAdding = "\nВыберите пересекаемые полилинии\n";
                pso_1.SingleOnly = false;
                pso_1.SinglePickInSpace = false;
                PromptSelectionResult result_1 = ed.GetSelection(pso_1, filter);
                if (result_1.Status == PromptStatus.OK) // в случае, если пользователь выбрал линии, то идем дальше
                {
                    SelectionSet intersectedSel = result_1.Value; // создаём переменную для записи в неё выбранного набора
                    using (Transaction Trans = db.TransactionManager.StartTransaction()) // начинаем транзакцию
                    {
                        try //начинаем обработку с блоком конструкцией улавливания ошибок
                        {
                            acBlkTbl = (BlockTable)Trans.GetObject(db.BlockTableId, OpenMode.ForRead, false, true);      //открываем для чтения класс BlockTable
                            acBlkTblRec = Trans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false, true) as BlockTableRecord;

                            Polyline polyline_Cutting = (Polyline)Trans.GetObject(cuttingSel.GetObjectIds()[0], OpenMode.ForRead, false, true);
                            if (polyline_Cutting != null)
                            {
                                //пока для проверки создаем вспомогательное сообщение
                             //   ed.WriteMessage($"Выбрана секущая полилиния с Id {polyline_Cutting.ObjectId} количеством вершин {polyline_Cutting.NumberOfVertices};\n");
                               
                                foreach (SelectedObject sObj in intersectedSel) //перебираем каждый выбранный объект (пересекаемую линию)
                                {
                                    Polyline polyline_Intersected = (Polyline)Trans.GetObject(sObj.ObjectId, OpenMode.ForRead, false, true);
                                    if (polyline_Intersected !=null)
                                    {
                                        //если выбраны полилинии, и нам вообще ничего не препятствует,
                                        //то ищем точки пересечения секущей и пересекаемой линий
                                        //пока для проверки создаем вспомогательное сообщение
                                      //  ed.WriteMessage($"Анализируем пересекаемую полилинию с Id {polyline_Intersected.ObjectId} количеством вершин {polyline_Intersected.NumberOfVertices};\n");
                                      added_Vertex_Polyline(polyline_Cutting, polyline_Intersected);
                                      //  ed.WriteMessage($"Функция \"added_Vertex_Polyline\" вроде отработала успешно\n");

                                    }
                                    else
                                    {
                                        ed.WriteMessage("Не найдено ObjectId пересекаемой полилинии\n");
                                        Trans.Abort();
                                    }

                                }//берем следующий объект из пересекаемого набора
                            }
                            else
                            {
                                ed.WriteMessage("Не найдено ObjectId секущей полилинии\n");
                                Trans.Abort();
                            }
                            Trans.Commit();
                        }
#if NCAD
                        catch (Teigha.Runtime.Exception ex)
#else
                        catch (Autodesk.AutoCAD.Runtime.Exception ex) 
#endif
                        {
                            ed.WriteMessage($"В процессе работы команды Test_Intersection обнаружена ошибка: {ex.Message}\n");
                            Trans.Abort();
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"В процессе работы команды Test_Intersection обнаружена ошибка: {ex.Message}\n");
                            Trans.Abort();
                        }


                    }

                }  else
                {
                    ed.WriteMessage("Не выбраны пересекаемые полилинии\n");
                    return;
                }

            }
            else
            {
                ed.WriteMessage("Не выбрана секущая полилиния\n");

            }
        }

        //________________________________________________________________________________________________________________
        [CommandMethod("Sdelay_PoperCS", CommandFlags.Modal | CommandFlags.UsePickSet | CommandFlags.Session)]
       // [Obsolete]
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
                            bool need_MP=false;
                            string textMP_style = "";
                            double textMP_height = 2;
                            string textMP_layer = "";
                            if (form1.checkBox1.Checked)
                            {
                                need_MP = true;
                                
                                textMP_layer = form1.textBox10.Text;
                                textMP_style= form1.textBox12.Text;
                                textMP_height=Convert.ToDouble(form1.textBox11.Text);
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
                                    //__________________________если пересекаемый объект - в слое 29 и фиолетового цвета - пропускаем его____________________________________
                                    if (popent.Layer == "29" && popent.ColorIndex == 6) continue;

                                        //______________________________________________________________________________________________________________
                                        Point3dCollection poper3dCol = new Point3dCollection();
                                        //находим точки пересечения сегмента полилинии разреза и характерной линии (может быть несколько)
                                      //  MyLine.IntersectWith(popent, Intersect.OnBothOperands, poper3dCol, 0, 0);
                                    MyLine.IntersectWith(popent, Intersect.OnBothOperands, poper3dCol, IntPtr.Zero, IntPtr.Zero);

                                    for (int j=0; j<= poper3dCol.Count - 1; j++)
                                        {
                                            H1 = Vychisli_Z(PointInVertCol[i], poper3dCol[j], PointInVertCol[i + 1]);
                                            //_________________________________________создаем точку пересечения линии поперечника с характерной линией_________________________________________________________________________________
                                            Point3d PopPointWithH = new Point3d(Math.Round(poper3dCol[j].X, 2), Math.Round(poper3dCol[j].Y, 2), Math.Round(PointInVertCol[i].Z + H1, 3));
                                            //______________________________________Здесь вставляем модуль вставки в чертеж описания точки____________________________________________________________________________________

                                            SozdanieKodaTochkiPopera(acBlkTblRec, PopPointWithH, popent, MyLine);
                                            SozdaniePodpisiTochkiPopera(acBlkTblRec, PopPointWithH, popent, MyLine);
                                        //--если в настройках нажат соответствующий чекбокс, то создаем тексты с междопутьями на плане---------------
                                        if (need_MP&& popent.Layer == "29")
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
                                if (Points_GR.Count>1)
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
                               
                            }else ed.WriteMessage("Процедура создания междопутий пропущена из-за настроек\n");


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
                                            if (countDel==0)
                                            {
                                                hasDubls=false;
                                            }
                                }
                                
                            }
                            //_________________________________________________________
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
            double S1 = line1.Length;
            double S2 = line2.Length;
            double Hfull = p3.Z - p1.Z;
            double Spol = S1 + S2;
            double TanA = Hfull / Spol;
                  double H1 = S1 * TanA;
            line1.Dispose();
            line2.Dispose();
           return H1;
        }
        public PromptEntityResult select_Entity(Type myType,String myMessage)
        {
            PromptEntityOptions promptPointoptions = new PromptEntityOptions(myMessage+"\n");
            promptPointoptions.SetRejectMessage($"Выбран объект не {myType}!\n");
            promptPointoptions.AllowNone = false;
            promptPointoptions.AddAllowedClass(myType, true);
            PromptEntityResult promptResult_1 = ed.GetEntity(promptPointoptions);
            if (promptResult_1.Status == PromptStatus.OK)
            {
                return promptResult_1;
            }
            else
            {
                return null;
            }
            
        }
        [CommandMethod("test_Type_Entity", CommandFlags.UsePickSet)]
        public void test_Type_Entity()
        {
            BlockTable acBlkTbl;   //объявляем переменные для базы с примитивами чертежа 
            BlockTableRecord acBlkTblRec;
            PromptEntityOptions promptPointoptions = new PromptEntityOptions("Выберите объект, тип которого хотите узнать\n");
            promptPointoptions.AllowNone = false;
            PromptEntityResult promptResult_1 = ed.GetEntity(promptPointoptions);
            if (promptResult_1.Status == PromptStatus.OK)
            {
                using (Transaction Trans = db.TransactionManager.StartTransaction())
                {
                    acBlkTbl = (BlockTable)Trans.GetObject(db.BlockTableId, OpenMode.ForRead, false, true);      //открываем для чтения класс BlockTable
                    acBlkTblRec = Trans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false, true) as BlockTableRecord;
                    Entity entity = (Entity)Trans.GetObject(promptResult_1.ObjectId, OpenMode.ForRead, false, true);
                    ed.WriteMessage($"Выбран объект типа: {entity.GetType().ToString()})\n");
                    switch (entity.GetType().ToString())
                    {
#if NCAD
                        case "Teigha.DatabaseServices.Face":
#else
                        case "Autodesk.AutoCAD.DatabaseServices.Face":
#endif
                            Face anyFace = (Face)Trans.GetObject(promptResult_1.ObjectId, OpenMode.ForRead, false, true);
                            Point3d point_1 = new Point3d(anyFace.GetVertexAt(0).X, anyFace.GetVertexAt(0).Y, anyFace.GetVertexAt(0).Z);
                            Point3d point_2 = new Point3d(anyFace.GetVertexAt(1).X, anyFace.GetVertexAt(1).Y, anyFace.GetVertexAt(1).Z);
                            Point3d point_3 = new Point3d(anyFace.GetVertexAt(2).X, anyFace.GetVertexAt(2).Y, anyFace.GetVertexAt(2).Z);
                            ed.WriteMessage($"Координаты вершин выбранной грани: {point_1}\t{point_2}\t{point_3}\n");
                            break;


                    }
                    Trans.Commit();
                }

            }
        }
        [CommandMethod("create_point3D_on_Face")]
        public void create_point3D_on_Face()
        {
            PromptPointOptions promptPointoptions = new PromptPointOptions("Укажите точку на экране, где хотите создать 3-д точку\n");
            promptPointoptions.AllowNone = false;
            PromptPointResult promptResult_1 = ed.GetPoint(promptPointoptions);
            if (promptResult_1.Status == PromptStatus.OK)
            {
                TypedValue[] TvFace = new TypedValue[1];
                TvFace.SetValue(new TypedValue((int)(DxfCode.Start), "3DFACE"), 0);
                SelectionFilter filterFace = new SelectionFilter(TvFace);
                // PromptSelectionResult resultFace = ed.SelectAll(filterFace);
                PromptSelectionResult resultFace =ed.SelectCrossingWindow(new Point3d(promptResult_1.Value.X+100, promptResult_1.Value.Y-100,0), new Point3d(promptResult_1.Value.X - 100, promptResult_1.Value.Y + 100, 0), filterFace);
                SelectionSet FaceSel = resultFace.Value;
                create_point_onFace(promptResult_1.Value, FaceSel);
            }
        }

        [CommandMethod("create_point3D_on_each_vertex", CommandFlags.Modal | CommandFlags.UsePickSet | CommandFlags.Session)]
        public void create_point3D_on_each_vertex()
        {
            BlockTable acBlkTbl;   //объявляем переменные для базы с примитивами чертежа 
            BlockTableRecord acBlkTblRec;
            PromptEntityResult result = select_Entity(typeof(Polyline), "Выберите полилинию, на вершинах которой надо создать точки\n");
             if (result!=null)
             {
                using (Transaction Trans = db.TransactionManager.StartTransaction())
                {
                    DocumentLock docklock = doc.LockDocument();
                    acBlkTbl = (BlockTable)Trans.GetObject(db.BlockTableId, OpenMode.ForRead, false, true);      //открываем для чтения класс BlockTable
                    acBlkTblRec = Trans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false, true) as BlockTableRecord;
                    Polyline myPoly=(Polyline)Trans.GetObject(result.ObjectId,OpenMode.ForRead, false, true);
                    //Point3d point_1 = myPoly.GeometricExtents.MaxPoint;
                    //Point3d point_2 = myPoly.GeometricExtents.MinPoint;
                    //Point3dCollection point3DCollectionPoly= new Point3dCollection();
                    //for (int i=0; i < myPoly.NumberOfVertices;++i)
                    //{
                    //    point3DCollectionPoly.Add(myPoly.GetPoint3dAt(i));
                    //}
                    //TypedValue[] TvFace = new TypedValue[1];
                    //TvFace.SetValue(new TypedValue((int)(DxfCode.Start), "3DFACE"), 0);
                    //SelectionFilter filterFace = new SelectionFilter(TvFace);
                    //// PromptSelectionResult resultFace = ed.SelectAll(filterFace);
                    //// PromptSelectionResult resultFace = ed.SelectCrossingWindow(new Point3d(anyPoint.X + 100, anyPoint.Y - 100, 0), new Point3d(anyPoint.X - 100, anyPoint.Y + 100, 0), filterFace);
                    //PromptSelectionResult resultFace = ed.SelectCrossingPolygon(point3DCollectionPoly, filterFace);
                    //if (resultFace.Status != PromptStatus.OK)
                    //{


                    //}
                    //if (resultFace.Status==PromptStatus.OK)
                    //{
                    //    SelectionSet FaceSel = resultFace.Value;
                    //    for (int i = 0; i < myPoly.NumberOfVertices; i++)
                    //    {
                    //        create_point_onFace(myPoly.GetPoint3dAt(i), FaceSel);
                    //    }
                    //    point3DCollectionPoly.Clear();
                    //    point3DCollectionPoly.Dispose();
                    //    Trans.Commit();
                    //    docklock.Dispose();
                    //}
                    //else
                    //{
                    //    ed.WriteMessage("Не получилось создать набор из граней\n");
                    //    Trans.Abort();
                    //}
                    for (int i = 0; i < myPoly.NumberOfVertices; i++)
                    {
                        Point3d anyPoint = myPoly.GetPoint3dAt(i);
                        TypedValue[] TvFace = new TypedValue[1];
                        TvFace.SetValue(new TypedValue((int)(DxfCode.Start), "3DFACE"), 0);
                        SelectionFilter filterFace = new SelectionFilter(TvFace);
                        // PromptSelectionResult resultFace = ed.SelectAll(filterFace);
                         PromptSelectionResult resultFace = ed.SelectCrossingWindow(new Point3d(anyPoint.X + 100, anyPoint.Y - 100, 0), new Point3d(anyPoint.X - 100, anyPoint.Y + 100, 0), filterFace);
                      //  PromptSelectionResult resultFace = ed.SelectCrossingWindow(point_2, new Point3d(anyPoint.X - 100, anyPoint.Y + 100, 0), filterFace);
                        if (resultFace == null)
                        {
                            ed.WriteMessage($"Для точки с координатами {anyPoint} не удалось выбрать грани\n");
                            return;
                        }
                        SelectionSet FaceSel = resultFace.Value;
                        create_point_onFace(anyPoint, FaceSel);
                    }
                    Trans.Commit();
                    docklock.Dispose();

                }

             }else
             {
                ed.WriteMessage("Не выбрана полилиния\n");
                return;
             }
        }

        public Point3d get_projection_on_rebro(Point3d anyPoint3D, Point3d point_1, Point3d point_2)
        {
            double myX = double.MinValue;
            double myY = double.MinValue;
            using (Line rebro=new Line (point_1, point_2))
            {
                Line b_line = new Line(point_1, anyPoint3D);
                double b = b_line.Length;
                double a=rebro.Length;
                Line c_line = new Line(point_2, anyPoint3D);
                double c = c_line.Length;
                double S = (a * a + b * b - c * c) / (2 * a);
                myX = point_1.X + ((S / a) * (point_2.X - point_1.X));
                myY = point_1.Y + ((S / a) * (point_2.Y - point_1.Y));
            }
            Point3d proj_point = new Point3d(myX,myY,anyPoint3D.Z);

            return proj_point;
        }
       public Point3d get_point_on_Face(Point3d anyPoint3D,Face anyFace)
       {
            Point3d proj_1 = get_projection_on_rebro(anyPoint3D, anyFace.GetVertexAt(0), anyFace.GetVertexAt(1));
            double dH_1 = Vychisli_Z(anyFace.GetVertexAt(0), proj_1, anyFace.GetVertexAt(1));
            double H1 = anyFace.GetVertexAt(0).Z + dH_1;
            Point3d proj_2 = get_projection_on_rebro(anyPoint3D, anyFace.GetVertexAt(1), anyFace.GetVertexAt(2));
            double dH_2 = Vychisli_Z(anyFace.GetVertexAt(1), proj_2, anyFace.GetVertexAt(2));
            double H2 = anyFace.GetVertexAt(1).Z + dH_2;
            Point3d proj_3 = get_projection_on_rebro(anyPoint3D, anyFace.GetVertexAt(2), anyFace.GetVertexAt(0));
            double dH_3 = Vychisli_Z(anyFace.GetVertexAt(2), proj_3, anyFace.GetVertexAt(0));
            double H3 = anyFace.GetVertexAt(2).Z + dH_3;
            double H_on_Face=(H1+ H2+ H3)/3;
            Point3d point_on_Face = new Point3d(anyPoint3D.X, anyPoint3D.Y,H_on_Face);
            return point_on_Face;
       }
        public Line point_on_rebro(Point3d anyPoint3D, Face anyFace)
        {
            Point3d vert_1 = new Point3d(anyFace.GetVertexAt(0).X, anyFace.GetVertexAt(0).Y,0);
            Point3d vert_2 = new Point3d(anyFace.GetVertexAt(1).X, anyFace.GetVertexAt(1).Y, 0);
            Point3d vert_3 = new Point3d(anyFace.GetVertexAt(2).X, anyFace.GetVertexAt(2).Y, 0);
            Line line_1= new Line(vert_1, vert_2);
            Line line_a = new Line(vert_1, new Point3d(anyPoint3D.X, anyPoint3D.Y, 0));
            Line line_b = new Line(vert_2, new Point3d(anyPoint3D.X, anyPoint3D.Y, 0));
            double my_Dist = line_a.Length+line_b.Length - line_1.Length;
            my_Dist = my_Dist/ line_1.Length;
                if (my_Dist < 0.002)
                {
                    Line line = new Line(anyFace.GetVertexAt(0), anyFace.GetVertexAt(1));
                    line_1.Dispose();
                    return line;
                }
                else
                {
                    line_1= new Line(vert_2, vert_3);
                    line_a = new Line(vert_2, new Point3d(anyPoint3D.X, anyPoint3D.Y, 0));
                    line_b = new Line(vert_3, new Point3d(anyPoint3D.X, anyPoint3D.Y, 0));
                    my_Dist =  line_a.Length + line_b.Length-line_1.Length;
                    my_Dist = my_Dist/ line_1.Length;
                    if (my_Dist < 0.002)
                    {
                     Line line = new Line(anyFace.GetVertexAt(1), anyFace.GetVertexAt(2));
                     line_1.Dispose();
                     line_a.Dispose();
                     line_b.Dispose();
                     return line;
                    }
                        else
                        {
                            line_1= new Line(vert_3, vert_1);
                            line_a = new Line(vert_3, new Point3d(anyPoint3D.X, anyPoint3D.Y, 0));
                            line_b = new Line(vert_1, new Point3d(anyPoint3D.X, anyPoint3D.Y, 0));
                            my_Dist = line_a.Length + line_b.Length - line_1.Length;
                            my_Dist = my_Dist / line_1.Length;
                            if (my_Dist < 0.002)
                            {
                                    Line line = new Line(anyFace.GetVertexAt(2), anyFace.GetVertexAt(0));
                                    line_1.Dispose();
                                line_a.Dispose();
                                line_b.Dispose();
                                return line;
                            }
                            else 
                            {
                                line_1.Dispose();
                                line_a.Dispose();
                                line_b.Dispose();
                               // ed.WriteMessage($"Точка лежит внутри грани\n");
                                return null;
                            }
                        }
                }
        }
        public bool point_inside_face(Point3d anyPoint3D,Face anyFace)
        {
            bool inside=false;
            Point3d point_1=new Point3d(anyFace.GetVertexAt(0).X, anyFace.GetVertexAt(0).Y,0);
            Point3d point_2 = new Point3d(anyFace.GetVertexAt(1).X, anyFace.GetVertexAt(1).Y,0);
            Point3d point_3 = new Point3d(anyFace.GetVertexAt(2).X, anyFace.GetVertexAt(2).Y, 0);

            if (point_1 != point_2 && point_2 != point_3)
            {
                Point3d med_point_1 = get_middle_point3D(point_2, point_3);
                Point3d med_point_2 = get_middle_point3D(point_1, point_3);
                Point3dCollection collection_mediana = new Point3dCollection();
                Line mediana_1= new Line(point_1, med_point_1);
                Line mediana_2 = new Line(point_2, med_point_2);
                mediana_1.IntersectWith(mediana_2, Intersect.OnBothOperands, collection_mediana, IntPtr.Zero, IntPtr.Zero);
                if (collection_mediana.Count>0)
                {
                    Point3d point_center_triangle = collection_mediana[0];
                    //ed.WriteMessage($"У грани получена центральная точка {point_center_triangle}\n");//----вспомогательно-----
                    Line Rasst_do_Point = new Line(point_center_triangle,new Point3d(anyPoint3D.X,anyPoint3D.Y,0));
                    Point3dCollection collection = new Point3dCollection();
                    Line line_1 = new Line(new Point3d(point_1.X, point_1.Y, 0), new Point3d(point_2.X, point_2.Y, 0));
                    Line line_2 = new Line(new Point3d(point_3.X, point_3.Y, 0), new Point3d(point_2.X, point_2.Y, 0));
                    Line line_3 = new Line(new Point3d(point_1.X, point_1.Y, 0), new Point3d(point_3.X, point_3.Y, 0));
                    Rasst_do_Point.IntersectWith(line_1, Intersect.OnBothOperands, collection, IntPtr.Zero, IntPtr.Zero);
                    if ((collection.Count == 0) || (Round(collection[0].X,3) ==Round(anyPoint3D.X,3) && Round(collection[0].Y,3) ==Round(anyPoint3D.Y,3)))
                    {
                        collection.Clear();
                        Rasst_do_Point.IntersectWith(line_2, Intersect.OnBothOperands, collection, IntPtr.Zero, IntPtr.Zero);
                        if ((collection.Count == 0) || (Round(collection[0].X, 3) == Round(anyPoint3D.X, 3) && Round(collection[0].Y, 3) == Round(anyPoint3D.Y, 3)))
                        {
                            collection.Clear();
                            Rasst_do_Point.IntersectWith(line_3, Intersect.OnBothOperands, collection, IntPtr.Zero, IntPtr.Zero);
                            if ((collection.Count == 0) || (Round(collection[0].X, 3) == Round(anyPoint3D.X, 3) && Round(collection[0].Y, 3) == Round(anyPoint3D.Y, 3)))
                            {
                                collection.Clear();
                               // ed.WriteMessage($"В работу ушла грань с центральной точкой {point_center_triangle}\n");
                                return true;
                            }
                            else
                            {
                                line_1.Dispose();
                                line_2.Dispose();
                                line_3.Dispose();
                                Rasst_do_Point.Dispose();
                                collection.Dispose();
                                return false;
                            }
                        }
                        else
                        {
                            line_1.Dispose();
                            line_2.Dispose();
                            line_3.Dispose();
                            Rasst_do_Point.Dispose();
                            collection.Dispose();
                            return false;
                        }
                    }
                    else
                    {
                        line_1.Dispose();
                        line_2.Dispose();
                        line_3.Dispose();
                        Rasst_do_Point.Dispose();
                        collection.Dispose();
                       // ed.WriteMessage($"Грань не подошла\n");
                        return false;
                    }


                }
                else
                {
                    ed.WriteMessage($"У грани с вершинами {point_1}\t{point_2}\t{point_3} не удалось найти медиану\n");//----вспомогательно-----
                }
                

            }

            return inside;
        }
        public void create_point_onFace(Point3d anyPoint3D, SelectionSet FaceSel)
        {
            BlockTable acBlkTbl;   //объявляем переменные для базы с примитивами чертежа 
            BlockTableRecord acBlkTblRec;
            if (FaceSel !=null)
            {
               // ed.WriteMessage($"В работу попало {FaceSel.Count} граней\n");
                using (Transaction Trans = db.TransactionManager.StartTransaction())
                {
                    int count_made = 0;
                    acBlkTbl = (BlockTable)Trans.GetObject(db.BlockTableId, OpenMode.ForRead, false, true);      //открываем для чтения класс BlockTable
                    acBlkTblRec = Trans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false, true) as BlockTableRecord;
                    foreach (SelectedObject sObj in FaceSel)
                    {
                        
                        Face anyFace = (Face)Trans.GetObject(sObj.ObjectId, OpenMode.ForRead, false, true);
                        if (anyFace != null)
                        {
                            bool check_point = point_inside_face(anyPoint3D, anyFace);
                          //  anyFace.UpgradeOpen();//-------вспомогательно--------
                            //anyFace.ColorIndex = 4;
                            //anyFace.LineWeight = LineWeight.LineWeight009;//-------вспомогательно--------
                            if (check_point)
                            {
                                // PlanarEntity my_plane = anyFace.GetPlane();// надо правильно находить плоскость
                                Plane my_plane=anyFace.GetPlane();
                                Vector3d vector_Z = Vector3d.ZAxis;
                                Point3d point3D=new Point3d();
                                Line line_rebro = point_on_rebro(anyPoint3D, anyFace);
                                if (line_rebro != null)
                                {
                                    
                                    double dH=Vychisli_Z(line_rebro.StartPoint,anyPoint3D,line_rebro.EndPoint);
                                    point3D = new Point3d(anyPoint3D.X, anyPoint3D.Y, line_rebro.StartPoint.Z + dH);
                                    if (point3D == null) ed.WriteMessage("Не сработал метод Vychisli_Z\n");
                                    line_rebro.Dispose();
                                    //acBlkTblRec.AppendEntity(line_rebro);
                                    //Trans.AddNewlyCreatedDBObject(line_rebro, true);
                                    //line_rebro.ColorIndex = 5;
                                    //line_rebro.LineWeight = LineWeight.LineWeight035;
                                    // ed.WriteMessage($"Точка создана на ребре {line_rebro.StartPoint} - {line_rebro.EndPoint}\n");
                                }
                                else
                                {
                                    point3D = anyPoint3D.Project(my_plane, vector_Z);
                                    if (point3D == null) ed.WriteMessage("Не сработал метод anyPoint3D.Project\n");
                                    // ed.WriteMessage($"Точка создана на грани, отметка {anyPoint3D.Z} \n");
                                    // anyFace.UpgradeOpen();
                                    //anyFace.ColorIndex = 1;
                                    //anyFace.LineWeight = LineWeight.LineWeight035;
                                }
                                DBPoint new_Point= new DBPoint(point3D);
                                ++count_made;
                                acBlkTblRec.AppendEntity(new_Point);
                                Trans.AddNewlyCreatedDBObject(new_Point, true);
                                Trans.Commit();
                                ed.WriteMessage($"Создана точка с координатами {new_Point.Position}\n");
                                my_plane.Dispose();
                                break;
                            }
                        }
                    }
                   // ed.WriteMessage($"Создано {count_made} точек\n");
                    if (count_made == 0) 
                    {
                        ed.WriteMessage($"Точка {anyPoint3D} не попала ни в одно ребро\n");
                        Trans.Abort();
                    } 
                   
                }
            }
            else
            {
                ed.WriteMessage("В чертеже не обнаружено 3-д граней\n");
                return;
            }
            
        }

        [CommandMethod("Interpol", CommandFlags.Modal | CommandFlags.UsePickSet | CommandFlags.Session)]
        public void Interpol()
        {
            BlockTable acBlkTbl;   //объявляем переменные для базы с примитивами чертежа 
            BlockTableRecord acBlkTblRec;
            Point3d point3D_1 = new Point3d();
            Point3d point3D_2 = new Point3d();
            Point3d point3D_3 = new Point3d();
            DocumentLock docklock = doc.LockDocument(DocumentLockMode.ExclusiveWrite,null,null,true);
            using (Transaction Trans = db.TransactionManager.StartTransaction()) // начинаем транзакцию
            {
                try
                {
                    acBlkTbl = (BlockTable)Trans.GetObject(db.BlockTableId, OpenMode.ForRead, false, true);      //открываем для чтения класс BlockTable
                    acBlkTblRec = Trans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false, true) as BlockTableRecord;
                    PromptEntityResult promptResult_1 = select_Entity(typeof(DBPoint), "Выберите первую точку");
                    if (promptResult_1 != null)
                    {
                        DBPoint dBPoint_1 = (DBPoint)Trans.GetObject(promptResult_1.ObjectId, OpenMode.ForRead,false);
                        point3D_1 = dBPoint_1.Position;
                    ed.WriteMessage($"Получена первая точка {dBPoint_1.Position}\n");
                        PromptEntityResult promptResult_2 = select_Entity(typeof(DBPoint), "Выберите вторую точку");
                        if (promptResult_2 != null)
                        {
                            DBPoint dBPoint_2 = (DBPoint)Trans.GetObject(promptResult_2.ObjectId, OpenMode.ForRead, false);
                            point3D_2 = dBPoint_2.Position;
                        ed.WriteMessage($"Получена вторая точка {dBPoint_2.Position}\n");
                        //Предлагаем выбрать полилинию, если пользователь откажется,
                        //то указываем вручную на экране, где создать 3-D точку
                        PromptEntityResult promptResult_3 = select_Entity(typeof(Polyline), "Выберите полилинию, на пересечении с которой построится точка");
                                if (promptResult_3 == null)
                                {
                                    PromptPointOptions pointOptions = new PromptPointOptions("Укажите, где Вы хотите создать точку\n");
                                    pointOptions.AllowNone = false;
                                    pointOptions.BasePoint = point3D_2;
                                    pointOptions.UseDashedLine = true;
                                    pointOptions.UseBasePoint = true;
                                    PromptPointResult pointResult = ed.GetPoint(pointOptions);
                                    if (pointResult.Status == PromptStatus.OK)
                                    {
                                        point3D_3 = pointResult.Value;
                                        double inter_H = Vychisli_Z(point3D_1, point3D_3, point3D_2); 
                                        ed.WriteMessage($"Получено превышение {inter_H}\n");
                                        Point3d inter_Point3D= new Point3d(point3D_3.X, point3D_3.Y, point3D_1.Z + inter_H);
                                        ed.WriteMessage($"Получена искомая точка {inter_Point3D}\n");
                                        DBPoint newPoint = new DBPoint(inter_Point3D);
                                        acBlkTblRec.AppendEntity(newPoint);
                                        Trans.AddNewlyCreatedDBObject(newPoint, true);
                                        Trans.Commit();
                                    }
                                    else
                                    {
                                        ed.WriteMessage("Не указана точка вставки\n");
                                        Trans.Abort();
                                    }

                                }
                                else 
                                {
                                    //если нам передана полилиния, то создаем точку
                                    //в виртуальном месте пересечения отрезка из 1 и 2 точек и полилинии
                                    Polyline sec_polyline = (Polyline)Trans.GetObject(promptResult_3.ObjectId, OpenMode.ForWrite, false);
                                    sec_polyline.Elevation = 0;
                                    using (Line virtual_line = new Line(new Point3d(point3D_1.X, point3D_1.Y, 0), new Point3d(point3D_2.X, point3D_2.Y, 0))) 
                                    {
                                        Point3dCollection intersect_col = new Point3dCollection();
                                        virtual_line.IntersectWith(sec_polyline, Intersect.OnBothOperands, intersect_col, IntPtr.Zero, IntPtr.Zero);
                                        if (intersect_col.Count > 0)
                                        {
                                            point3D_3 = intersect_col[0];
                                            double inter_H = Vychisli_Z(point3D_1, point3D_3, point3D_2);
                                            ed.WriteMessage($"Получено превышение {inter_H}\n");
                                            Point3d inter_Point3D = new Point3d(point3D_3.X, point3D_3.Y, point3D_1.Z+inter_H);
                                            DBPoint newPoint = new DBPoint(inter_Point3D);
                                            acBlkTblRec.AppendEntity(newPoint);
                                            Trans.AddNewlyCreatedDBObject(newPoint, true);
                                            //теперь вставляем вершину в выбранную полилинию в этом месте
                                            int insert_place = find_addvertex_index(sec_polyline, inter_Point3D);
                                            sec_polyline.AddVertexAt(insert_place, new Point2d(point3D_3.X, point3D_3.Y), 0, 0, 0);
                                            Trans.Commit();
                                        }
                                        else
                                        {
                                            ed.WriteMessage("Не найдены точки пересечения с линией\n");
                                            Trans.Abort();
                                        }
                                    }

                                    
                                }
  
                        }
                        else
                        {
                            ed.WriteMessage("Не выбрана вторая точка\n");
                            Trans.Abort();
                        }
                    }
                    else
                    {
                        ed.WriteMessage("Не выбрана первая точка\n");
                        Trans.Abort();
                    }



                }
#if NCAD
                catch (Teigha.Runtime.Exception ex)
#else
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
#endif
                {
                    ed.WriteMessage($"В процессе работы команды \"Interpol\" обнаружена ошибка: {ex.Message}\n");
                    Trans.Abort();
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"В процессе работы команды \"Interpol\" обнаружена ошибка: {ex.Message}\n");
                    Trans.Abort();
                }
            }
            docklock.Dispose();
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
            bool myBool = false;
            TypedValue[] TvOpisanie = new TypedValue[2];
            TvOpisanie.SetValue(new TypedValue((int)(DxfCode.Start), "TEXT"), 0);
            TvOpisanie.SetValue(new TypedValue((int)(DxfCode.LayerName), "подписи поперечников"), 1);
            SelectionFilter filterOpisanie = new SelectionFilter(TvOpisanie);
            PromptSelectionResult resultOpisanie = ed.SelectAll(filterOpisanie);
            if (resultOpisanie.Status == PromptStatus.OK)
            {
                SelectionSet poperOpisanie = resultOpisanie.Value;
                using (Transaction Trans = db.TransactionManager.StartTransaction())
                {
                    foreach (SelectedObject ObjOpisanie in poperOpisanie) //ошибка здесь
                    {
                        DBText TryOpisanie = Trans.GetObject(ObjOpisanie.ObjectId, OpenMode.ForRead, false, false) as DBText;
                       // ed.WriteMessage($"Сравниваются координаты текста описания: {TryOpisanie.Position}\n и точки {myPoint3d}\n");
                        if ((Round(TryOpisanie.Position.X,2) == Round(myPoint3d.X,2)) & (Round(TryOpisanie.Position.Y,2) ==Round(myPoint3d.Y,2)))
                        {
                            myBool = true;
                            break;
                        } //End If
                        else
                        {
                            myBool = false;
                        }
                        
                    } 
                } //End Using
                  // 
            }
            else ed.WriteMessage("При выполнении функции ProveritNalichiePodpisiInTochka не найдены подписи поперечников\n");
            
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

        public static double Vychisli_S(Point3d p1, Point3d p2)
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
#if NCAD
                    case "Teigha.DatabaseServices.Polyline":
#else
                    case "Autodesk.AutoCAD.DatabaseServices.Polyline":
#endif
                        using ( Polyline ProverkaPolyLine = Trans.GetObject(popent.ObjectId, OpenMode.ForRead) as Polyline)
                        {
                            switch (ProverkaPolyLine.Linetype)
                            {
                                case "atp_122_v":
                                    OpisanieTochkiPopera = "водопровод";
                                    break;
                                case "atp_121_v":
                                    OpisanieTochkiPopera = "водопровод";
                                    break;
                                case "atp_122_d":
                                    OpisanieTochkiPopera = "дренаж";
                                    break;
                                case "atp_121_d":
                                    OpisanieTochkiPopera = "дренаж";
                                    break;
                                case "atp_122_k":
                                    OpisanieTochkiPopera = "канализация";
                                    break;
                                case "atp_121_k":
                                    OpisanieTochkiPopera = "канализация";
                                    break;
                                case "atp_122_kn":
                                    OpisanieTochkiPopera = "канализация напорная";
                                    break;
                                case "atp_122_kl":
                                    OpisanieTochkiPopera = "канализация ливневая";
                                    break;
                                case "atp_122_g":
                                    OpisanieTochkiPopera = "газопровод";
                                    break;
                                case "atp_121_g":
                                    OpisanieTochkiPopera = "газопровод";
                                    break;
                                case "atp_122_t":
                                    OpisanieTochkiPopera = "теплотрасса";
                                    break;
                                case "atp_121_t":
                                    OpisanieTochkiPopera = "теплотрасса";
                                    break;
                                case "atp_119_3":
                                    switch (ProverkaPolyLine.Layer)
                                    {
                                        case "7":
                                            OpisanieTochkiPopera = "кабель 0.4кВ";
                                            break;
                                        case "10":
                                            OpisanieTochkiPopera = "кабель СЦБ";
                                            break;
                                    } //End Select
                                    break;
                                case "atp_121_2":
                                    OpisanieTochkiPopera = "кабель 0.4кВ";
                                    break;
                                case "atp_120_3b":
                                    OpisanieTochkiPopera = "кабель 0.4кВ";
                                    break;
                                case "atp_119_1":
                                    OpisanieTochkiPopera = "кабель 10кВ";
                                    break;
                                case "atp_121_1":
                                    OpisanieTochkiPopera = "кабель 10кВ";
                                    break;
                                case "atp_120_3a":
                                    OpisanieTochkiPopera = "кабель 10кВ";
                                    break;
                                case "atp_133_t":
                                    OpisanieTochkiPopera = "кабель связи";
                                    break;
                                case "atp_133":
                                    OpisanieTochkiPopera = "кабель связи";
                                    break;
                                case "atp_121_3":
                                    OpisanieTochkiPopera = "кабель связи";
                                    break;
                                case "atp_122_vx":
                                    OpisanieTochkiPopera = "воздухопровод";
                                    break;
                                case "atp_121_vx":
                                    OpisanieTochkiPopera = "воздухопровод";
                                    break;
                                case "atp_5x2":
                                    OpisanieTochkiPopera = "край грунтовой дороги";
                                    break;
                                case "atp_472":
                                    OpisanieTochkiPopera = "железобетонный забор";
                                    break;
                                case "atp_473":
                                    OpisanieTochkiPopera = "железобетонный забор";
                                    break;
                                case "atp_474_2a":
                                    OpisanieTochkiPopera = "металлический забор";
                                    break;
                                case "atp_474_1b":
                                    OpisanieTochkiPopera = "металлический забор";
                                    break;
                                case "atp_475_1":
                                    OpisanieTochkiPopera = "деревянный забор";
                                    break;
                                case "atp_476_3":
                                    OpisanieTochkiPopera = "забор сетка-рабица";
                                    break;
                                case "atp_476_1":
                                    OpisanieTochkiPopera = "колючая проволока";
                                    break;
                                case "atp_477":
                                    OpisanieTochkiPopera = "изгородь";
                                    break;
                                case "atp_476_2a":
                                    OpisanieTochkiPopera = "гладкая проволока";
                                    break;
                                case "atp_475_4a":
                                    OpisanieTochkiPopera = "деревянный с опорами";
                                    break;
                                case "atp_475_2":
                                    OpisanieTochkiPopera = "деревянный решетчатый";
                                    break;
                                case "atp_278_6b":
                                    OpisanieTochkiPopera = "металлический парапет";
                                    break;
                                case "atp_278_5":
                                    OpisanieTochkiPopera = "каменный парапет";
                                    break;
                                case "atp_1.5x1.5":
                                    switch (ProverkaPolyLine.Layer)
                                    {
                                        case "20":
                                            OpisanieTochkiPopera = "край отмостки";
                                            break;
                                        case "23":
                                        OpisanieTochkiPopera = "край асфальта";
                                            break;
                                    }
                                    break;

                                case "Continuous":
                                switch (ProverkaPolyLine.Layer)
                                    {
                                        case "20":
                                            OpisanieTochkiPopera = "стена здания";
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
                                        case 256:
                                            OpisanieTochkiPopera = "Хпр.ХкВ";
                                            break;
                                        case 1:
                                            OpisanieTochkiPopera = "Хпр.ХкВ";
                                            break;
                                    }
                                    break;

                                case "12":
                                switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case int i when i != 7:
                                            OpisanieTochkiPopera = "смотри на плане";
                                            break;
                                    }
                                    break;
                                case "14":
                                    OpisanieTochkiPopera = "трубопровод спец.назначения";
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
#if NCAD
                    case "Teigha.DatabaseServices.Line":
#else

                    case "Autodesk.AutoCAD.DatabaseServices.Line":
#endif
                        using (Line ProverkaPolyLine = Trans.GetObject(popent.ObjectId, OpenMode.ForRead) as Line)
                        {
                            switch (ProverkaPolyLine.Linetype)
                            {
                                case "atp_122_v":
                                    OpisanieTochkiPopera = "водопровод";
                                    break;
                                case "atp_121_v":
                                    OpisanieTochkiPopera = "водопровод";
                                    break;
                                case "atp_122_d":
                                    OpisanieTochkiPopera = "дренаж";
                                    break;
                                case "atp_121_d":
                                    OpisanieTochkiPopera = "дренаж";
                                    break;
                                case "atp_122_k":
                                    OpisanieTochkiPopera = "канализация";
                                    break;
                                case "atp_121_k":
                                    OpisanieTochkiPopera = "канализация";
                                    break;
                                case "atp_122_kn":
                                    OpisanieTochkiPopera = "канализация напорная";
                                    break;
                                case "atp_122_kl":
                                    OpisanieTochkiPopera = "канализация ливневая";
                                    break;
                                case "atp_122_g":
                                    OpisanieTochkiPopera = "газопровод";
                                    break;
                                case "atp_121_g":
                                    OpisanieTochkiPopera = "газопровод";
                                    break;
                                case "atp_122_t":
                                    OpisanieTochkiPopera = "теплотрасса";
                                    break;
                                case "atp_121_t":
                                    OpisanieTochkiPopera = "теплотрасса";
                                    break;
                                case "atp_119_3":
                                    switch (ProverkaPolyLine.Layer)
                                    {
                                        case "7":
                                            OpisanieTochkiPopera = "кабель 0.4кВ";
                                            break;
                                        case "10":
                                            OpisanieTochkiPopera = "кабель СЦБ";
                                            break;
                                    } //End Select
                                    break;
                                case "atp_121_2":
                                    OpisanieTochkiPopera = "кабель 0.4кВ";
                                    break;
                                case "atp_120_3b":
                                    OpisanieTochkiPopera = "кабель 0.4кВ";
                                    break;
                                case "atp_119_1":
                                    OpisanieTochkiPopera = "кабель 10кВ";
                                    break;
                                case "atp_121_1":
                                    OpisanieTochkiPopera = "кабель 10кВ";
                                    break;
                                case "atp_120_3a":
                                    OpisanieTochkiPopera = "кабель 10кВ";
                                    break;
                                case "atp_133_t":
                                    OpisanieTochkiPopera = "кабель связи";
                                    break;
                                case "atp_133":
                                    OpisanieTochkiPopera = "кабель связи";
                                    break;
                                case "atp_121_3":
                                    OpisanieTochkiPopera = "кабель связи";
                                    break;
                                case "atp_122_vx":
                                    OpisanieTochkiPopera = "воздухопровод";
                                    break;
                                case "atp_121_vx":
                                    OpisanieTochkiPopera = "воздухопровод";
                                    break;
                                case "atp_5x2":
                                    OpisanieTochkiPopera = "край грунтовой дороги";
                                    break;
                                case "atp_472":
                                    OpisanieTochkiPopera = "железобетонный забор";
                                    break;
                                case "atp_473":
                                    OpisanieTochkiPopera = "железобетонный забор";
                                    break;
                                case "atp_474_2a":
                                    OpisanieTochkiPopera = "металлический забор";
                                    break;
                                case "atp_474_1b":
                                    OpisanieTochkiPopera = "металлический забор";
                                    break;
                                case "atp_475_1":
                                    OpisanieTochkiPopera = "деревянный забор";
                                    break;
                                case "atp_476_3":
                                    OpisanieTochkiPopera = "забор сетка-рабица";
                                    break;
                                case "atp_476_1":
                                    OpisanieTochkiPopera = "колючая проволока";
                                    break;
                                case "atp_477":
                                    OpisanieTochkiPopera = "изгородь";
                                    break;
                                case "atp_476_2a":
                                    OpisanieTochkiPopera = "гладкая проволока";
                                    break;
                                case "atp_475_4a":
                                    OpisanieTochkiPopera = "деревянный с опорами";
                                    break;
                                case "atp_475_2":
                                    OpisanieTochkiPopera = "деревянный решетчатый";
                                    break;
                                case "atp_278_6b":
                                    OpisanieTochkiPopera = "металлический парапет";
                                    break;
                                case "atp_278_5":
                                    OpisanieTochkiPopera = "каменный парапет";
                                    break;
                                case "atp_1.5x1.5":
                                    switch (ProverkaPolyLine.Layer)
                                    {
                                        case "20":
                                            OpisanieTochkiPopera = "край отмостки";
                                            break;
                                        case "23":
                                            OpisanieTochkiPopera = "край асфальта";
                                            break;
                                    }
                                    break;

                                case "Continuous":
                                    switch (ProverkaPolyLine.Layer)
                                    {
                                        case "20":
                                            OpisanieTochkiPopera = "стена здания";
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
                                        case 1:
                                            OpisanieTochkiPopera = "Хпр.ХкВ";
                                            break;
                                    }
                                    break;

                                case "12":
                                    switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case int i when i != 7:
                                            OpisanieTochkiPopera = "смотри на плане";
                                            break;
                                    }
                                    break;
                                case "14":
                                    OpisanieTochkiPopera = "трубопровод спец.назначения";
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
#if NCAD
                    case "Teigha.DatabaseServices.Spline":
#else
                    case "Autodesk.AutoCAD.DatabaseServices.Spline":
#endif
                        using (Spline ProverkaPolyLine = Trans.GetObject(popent.ObjectId, OpenMode.ForRead) as Spline)
                        {
                            switch (ProverkaPolyLine.Linetype)
                            {
                                case "atp_122_v":
                                    OpisanieTochkiPopera = "водопровод";
                                    break;
                                case "atp_121_v":
                                    OpisanieTochkiPopera = "водопровод";
                                    break;
                                case "atp_122_d":
                                    OpisanieTochkiPopera = "дренаж";
                                    break;
                                case "atp_121_d":
                                    OpisanieTochkiPopera = "дренаж";
                                    break;
                                case "atp_122_k":
                                    OpisanieTochkiPopera = "канализация";
                                    break;
                                case "atp_121_k":
                                    OpisanieTochkiPopera = "канализация";
                                    break;
                                case "atp_122_kn":
                                    OpisanieTochkiPopera = "канализация напорная";
                                    break;
                                case "atp_122_kl":
                                    OpisanieTochkiPopera = "канализация ливневая";
                                    break;
                                case "atp_122_g":
                                    OpisanieTochkiPopera = "газопровод";
                                    break;
                                case "atp_121_g":
                                    OpisanieTochkiPopera = "газопровод";
                                    break;
                                case "atp_122_t":
                                    OpisanieTochkiPopera = "теплотрасса";
                                    break;
                                case "atp_121_t":
                                    OpisanieTochkiPopera = "теплотрасса";
                                    break;
                                case "atp_119_3":
                                    switch (ProverkaPolyLine.Layer)
                                    {
                                        case "7":
                                            OpisanieTochkiPopera = "кабель 0.4кВ";
                                            break;
                                        case "10":
                                            OpisanieTochkiPopera = "кабель СЦБ";
                                            break;
                                    } //End Select
                                    break;
                                case "atp_121_2":
                                    OpisanieTochkiPopera = "кабель 0.4кВ";
                                    break;
                                case "atp_120_3b":
                                    OpisanieTochkiPopera = "кабель 0.4кВ";
                                    break;
                                case "atp_119_1":
                                    OpisanieTochkiPopera = "кабель 10кВ";
                                    break;
                                case "atp_121_1":
                                    OpisanieTochkiPopera = "кабель 10кВ";
                                    break;
                                case "atp_120_3a":
                                    OpisanieTochkiPopera = "кабель 10кВ";
                                    break;
                                case "atp_133_t":
                                    OpisanieTochkiPopera = "кабель связи";
                                    break;
                                case "atp_133":
                                    OpisanieTochkiPopera = "кабель связи";
                                    break;
                                case "atp_121_3":
                                    OpisanieTochkiPopera = "кабель связи";
                                    break;
                                case "atp_122_vx":
                                    OpisanieTochkiPopera = "воздухопровод";
                                    break;
                                case "atp_121_vx":
                                    OpisanieTochkiPopera = "воздухопровод";
                                    break;
                                case "atp_5x2":
                                    OpisanieTochkiPopera = "край грунтовой дороги";
                                    break;
                                case "atp_472":
                                    OpisanieTochkiPopera = "железобетонный забор";
                                    break;
                                case "atp_473":
                                    OpisanieTochkiPopera = "железобетонный забор";
                                    break;
                                case "atp_474_2a":
                                    OpisanieTochkiPopera = "металлический забор";
                                    break;
                                case "atp_474_1b":
                                    OpisanieTochkiPopera = "металлический забор";
                                    break;
                                case "atp_475_1":
                                    OpisanieTochkiPopera = "деревянный забор";
                                    break;
                                case "atp_476_3":
                                    OpisanieTochkiPopera = "забор сетка-рабица";
                                    break;
                                case "atp_476_1":
                                    OpisanieTochkiPopera = "колючая проволока";
                                    break;
                                case "atp_477":
                                    OpisanieTochkiPopera = "изгородь";
                                    break;
                                case "atp_476_2a":
                                    OpisanieTochkiPopera = "гладкая проволока";
                                    break;
                                case "atp_475_4a":
                                    OpisanieTochkiPopera = "деревянный с опорами";
                                    break;
                                case "atp_475_2":
                                    OpisanieTochkiPopera = "деревянный решетчатый";
                                    break;
                                case "atp_278_6b":
                                    OpisanieTochkiPopera = "металлический парапет";
                                    break;
                                case "atp_278_5":
                                    OpisanieTochkiPopera = "каменный парапет";
                                    break;
                                case "atp_1.5x1.5":
                                    switch (ProverkaPolyLine.Layer)
                                    {
                                        case "20":
                                            OpisanieTochkiPopera = "край отмостки";
                                            break;
                                        case "23":
                                            OpisanieTochkiPopera = "край асфальта";
                                            break;
                                    }
                                    break;

                                case "Continuous":
                                    switch (ProverkaPolyLine.Layer)
                                    {
                                        case "20":
                                            OpisanieTochkiPopera = "стена здания";
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
                                        case 1:
                                            OpisanieTochkiPopera = "Хпр.ХкВ";
                                            break;
                                    }
                                    break;

                                case "12":
                                    switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case int i when i != 7:
                                            OpisanieTochkiPopera = "смотри на плане";
                                            break;
                                    }
                                    break;
                                case "14":
                                    OpisanieTochkiPopera = "трубопровод спец.назначения";
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
#if NCAD
                        case "Teigha.DatabaseServices.Line":
#else
                        case "Autodesk.AutoCAD.DatabaseServices.Line":
#endif
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
#if NCAD
                        case "Teigha.DatabaseServices.Polyline":
#else
                        case "Autodesk.AutoCAD.DatabaseServices.Polyline":
#endif
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
#if NCAD
                catch (Teigha.Runtime.Exception ex)
#else
                catch (Autodesk.AutoCAD.Runtime.Exception ex )
#endif
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
                double HforVX= Convert.ToDouble(form1.textBox9.Text);
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

                            PromptStringOptions popFilePromptStringOptions=new PromptStringOptions("\nВведите имя поперечника\n");
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
                            bool popal=false;
                            Point3d ZeroPoint = new Point3d();
                            do
                            {
                                PromptPointOptions ppo = new PromptPointOptions("\nУкажите нулевую точку на линии поперечника\n");
                                ppo.AllowNone = false;
                                PromptPointResult ZeroPointResult = ed.GetPoint(ppo);
                                ZeroPoint = ZeroPointResult.Value; //пользователь задает нулевую точку отсчета на поперечника (от нее "лево" и "право")
                                popal = MyCommonFunctions.point_is_vertex(ZeroPoint, MyPl3d,Trans);
                                if (popal==false)
                                {
                                    ed.WriteMessage("Вы не попали в вершину рабочей линии поперечника!\n");
                                }
                            } while (popal == false);
                           
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
                                            if (Opisanie.StartsWith("водопровод") ==true || Opisanie.StartsWith("канализация") ==true || Opisanie.StartsWith("теплотрасса") == true || Opisanie.StartsWith("газопровод") == true)
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
                                    text_zero_rasst.TextString = "L="+ RasstDoZero.ToString("0.00", CultureInfo.InvariantCulture);
                                    text_zero_rasst.Layer = "расстояния до Zero";
                                    text_zero_rasst.Height = 0.2;
                                    text_zero_rasst.Rotation= 0;
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
#if NCAD
            catch (Teigha.Runtime.Exception ex)
#else
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
#endif
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
#if NCAD
                    case "Teigha.DatabaseServices.Polyline3d":
#else
                    case "Autodesk.AutoCAD.DatabaseServices.Polyline3d":
#endif
                        Polyline3d MyPl3d = Trans.GetObject(ent.ObjectId, OpenMode.ForWrite) as Polyline3d;
                double lomanaya = Vychisli_LomDlinu(MyPl3d, myPoint, zeroPoint);
                ed.WriteMessage($"\nДлина ломаной по полилинии между точками составила {lomanaya}");
                        break;
#if NCAD
                    case "Teigha.DatabaseServices.Polyline":
#else
                    case "Autodesk.AutoCAD.DatabaseServices.Polyline":
#endif
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
        [CommandMethod("Test_Risui_Poper", CommandFlags.Modal)]
        public static void Test_Risui_Poper()
        {
            //Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            //Editor ed = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor;
            //Database db = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
#if NCAD
            Document doc = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            //   Editor ed= Teigha.Editor.CommandContext.Editor;
            Database db = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database;
#else
        Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
        Editor ed = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor;
        Database db = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
#endif
            BlockTable acBlkTbl;   //объявляем переменные для базы с примитивами чертежа 
            BlockTableRecord acBlkTblRec;
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

                    string poperCaption = lines[0].Substring(5,lines[0].Length-5);
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
                    ObjectIdCollection colIdOpisanie=new ObjectIdCollection();
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

                            if (lines[i][0]=='1'|| lines[i][0] == '0' || lines[i][0] == '2' || lines[i][0] == '3' || lines[i][0] == '4' || lines[i][0] == '5' || lines[i][0] == '6' || lines[i][0] == '7' || lines[i][0] == '8' || lines[i][0] == '9')//если строка начинается с символа 1 или 0 (то есть код есть), то разбиваем массив на подстроки
                            {
                                string[] znach;
                                znach = lines[i].Split(new char[] { ' ' },4,StringSplitOptions.RemoveEmptyEntries);//разбиваем строку на подстроки по разделитель "пробел",
                                                                                                                   //максимум 4 подстроки (код,расстояние, высота, описание),
                                                                                                                   //последовательные пробелы исключаются
                        
                                    if (znach.Length > 1) 
                                    { 
                                                
                                            
                                            double Code = Convert.ToDouble(znach[0]);
                                            double deltaX = Round(Convert.ToDouble(znach[1]),2);//меняем знак расстояния в зависимости от "лево" или "право"
                                                if (deltaX == 0)
                                                {
                                                    changeZnak=false;
                                                }
                                                if (changeZnak == true)
                                                {
                                                    deltaX = -deltaX;
                                                }
                                            double deltaY = Round(Convert.ToDouble(znach[2]),2);
                                               if (Code == 0 || Code == 1|| Code==501)
                                                {
                                                    minY = Min(minY, startPoint3D.Y + deltaY * vertScaleKoef);
                                                    maxY=Max(maxY, startPoint3D.Y + deltaY * vertScaleKoef);
                                                }
                           
                                             string opisanie = "";
                                                if (znach.Length> 3)
                                                {
                                                    opisanie = znach[3];
                                                }
                                                //ed.WriteMessage($"\nСчитано следующее содержимое поперечника:\n {Code}  {deltaX}  {deltaY} {opisanie}\n");

                                                //в случае, если код точки - не подземка и не ИССО, то в линию поперечника добавляем вершину_________________
                            
                                                if (Code == 0||Code==2 || Code == 3 || Code == 4 || Code == 5 || Code == 6 || Code == 7 || Code == 8 || Code == 9 || Code == 10 || Code == 11 || Code == 15)
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
                                                        String roundDist = Abs(deltaX).ToString("0.00", CultureInfo.InvariantCulture);
                                                        newOpisanie.TextString = roundDist + " " + newOpisanie.TextString;
                                                        // newOpisanie.TextString = Convert.ToString(Abs(deltaX)) + " " + newOpisanie.TextString;
                                                    } else if (Code == 1)
                                                    {
                                                        newOpisanie.TextString = "путь " + newOpisanie.TextString;
                                                    }
                                                    // newOpisanie.Height = 2;
                                                    newOpisanie.Height = Convert.ToDouble(form1.textBox3.Text);
                                                    newOpisanie.Rotation = 1.5708;
                                                    newOpisanie.Position = new Point3d(startPoint3D.X + deltaX * horizScaleKoef, startPoint3D.Y + deltaY * vertScaleKoef, 0);
                                                     //dBTextsOpisanie.Append(newOpisanie);
                                                   if (newOpisanie.TextStyleName=="ATP")
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
                                                    newOtmetka.TextString = Convert.ToString(Round(deltaY,2));
                                                    newOtmetka.TextString = deltaY.ToString("0.00", CultureInfo.InvariantCulture);
                                                    if (Code == 1) 
                                                    {
                                                        newOtmetka.TextString = "СГР " + newOtmetka.TextString;
                                                    }
                                                    newOtmetka.Height = Convert.ToDouble(form1.textBox3.Text);
                                                    newOtmetka.Rotation = 1.5708;
                                                    newOtmetka.ColorIndex = 1;
                                                    newOtmetka.Position = new Point3d(startPoint3D.X + deltaX * horizScaleKoef, startPoint3D.Y -50 + deltaY * vertScaleKoef, 0);
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
                                                    if (Code == 121 || Code == 131 || Code == 141||Code==501)
                                                    {
                                                        Point3d absPodzemkaPt = new Point3d(startPoint3D.X + deltaX * horizScaleKoef, startPoint3D.Y + deltaY * vertScaleKoef,0);
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
                                                if (Code ==501) //в отдельную коллекцию собираем расстояния до точек с кодом ИССО. 
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
                                
                                                    newCode.Position = new Point3d(startPoint3D.X + deltaX * horizScaleKoef, startPoint3D.Y-70 + deltaY  * vertScaleKoef, 0);
                                                // dBTextsCodes.Append(newCode);
                                                    colIdCodes.Add(newCode.ObjectId);
                                           // ed.WriteMessage($"Считаны следующие значения: Код: {Code};\t расстояние: {deltaX};\tОтметка: {deltaY};\t Пояснение: {opisanie}\n");
                                    }
                            }



                        
                    }

                    //ed.WriteMessage($"Вычислено значение переменной MinY={minY}\n");
                   // ed.WriteMessage($"Установлено значение MinY: {minY};\n начальная точки по Х: {profLine.StartPoint.X}\n");
                    MyCommonFunctions.ChertiPoper(profLine, minY, maxY,colIdOpisanie, colIdOtmetki, colIdRasst, colIdCodes, Trans, acBlkTblRec, horizScaleKoef, vertScaleKoef, colIdPodzemkaRasst, colIdPodzemkaAbsolut, colIdRasstIsso, db.Textstyle, poperCaption, startPoint3D);
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
        [CommandMethod("ShowPoperTools", CommandFlags.Modal)]
        
        public void ShowPoperTools()
        {
           // Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(form1);
            Form1 form1 = new Form1();
#if NCAD
            HostMgd.ApplicationServices.Application.ShowModelessDialog(form1);
#else
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(form1);
#endif
            form1.Visible= true;

            //form1.Show();
            //form1.ShowDialog();
            //form1.ControlBox = true;

            //form1.Activate();
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
#if NCAD
            HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("Sdelay_PoperCS" + " ", false, false, true);
#else
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("Sdelay_PoperCS" + " ", false, false, true);
#endif
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
#if NCAD
            HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("PrintPopFileCS" + " ", false, false, true);
#else
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("PrintPopFileCS" + " ", false, false, true);
#endif
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
#if NCAD
           HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("DlinaLomanoi" + " ", false, false, true);
#else
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("DlinaLomanoi" + " ", false, false, true);
#endif
        }
    }

    public class CommandHandler_ribbonButton4 : System.Windows.Input.ICommand
    {
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object param)
        {
            return true;
        }

        public void Execute(object parameter)
        {
#if NCAD
            HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("Test_Risui_Poper" + " ", false, false, true);
#else

            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("Test_Risui_Poper" + " ", false, false, true);
#endif
        }
    }

    public class CommandHandler_ribbonButton5 : System.Windows.Input.ICommand
    {
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object param)
        {
            return true;
        }

        public void Execute(object parameter)
        {
#if NCAD
            HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("ShowPoperTools" + " ", false, false, true);
#else
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("ShowPoperTools" + " ", false, false, true);
#endif
        }
    }

    public class CommandHandler_ribbonButton6 : System.Windows.Input.ICommand
    {
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object param)
        {
            return true;
        }

        public void Execute(object parameter)
        {
#if NCAD
           HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("Sdelay_Zero" + " ", false, false, true);
#else
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("Sdelay_Zero" + " ", false, false, true);
#endif
        }
    }

    public class CommandHandler_ribbonButton7 : System.Windows.Input.ICommand
    {
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object param)
        {
            return true;
        }

        public void Execute(object parameter)
        {
#if NCAD
            HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("Make_Description" + " ", false, false, true);
#else
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("Make_Description" + " ", false, false, true);
#endif
        }
    }

    public class CommandHandler_ribbonButton8 : System.Windows.Input.ICommand
    {
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object param)
        {
            return true;
        }

        public void Execute(object parameter)
        {
#if NCAD
           HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("Test_Intersection" + " ", false, false, true);
#else
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("Test_Intersection" + " ", false, false, true);
#endif
        }
    }
    public class CommandHandler_ribbonButton9 : System.Windows.Input.ICommand
    {
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object param)
        {
            return true;
        }

        public void Execute(object parameter)
        {
#if NCAD
            HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("Interpol" + " ", false, false, true);
#else
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("Interpol" + " ", false, false, true);
#endif
        }
    }
    public class CommandHandler_ribbonButton10 : System.Windows.Input.ICommand
    {
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object param)
        {
            return true;
        }

        public void Execute(object parameter)
        {
#if NCAD
            HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("group_intersection" + " ", false, false, true);
#else
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("group_intersection" + " ", false, false, true);
#endif
        }
    }
    public class CommandHandler_ribbonButton11 : System.Windows.Input.ICommand
    {
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object param)
        {
            return true;
        }

        public void Execute(object parameter)
        {
#if NCAD
            HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("create_point3D_on_each_vertex" + " ", false, false, true);
#else
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("create_point3D_on_each_vertex" + " ", false, false, true);
#endif
        }
    }
}
       



    



           
       










