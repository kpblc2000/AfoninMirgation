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
        //делаем функцию приведения всех характерных линий на отметку 0

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
                need_DBPoint = true;
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
                        DBPoint selected_DBPoint = (DBPoint)Trans.GetObject(Sobj.ObjectId, OpenMode.ForRead, false);
                        if ((selected_DBPoint.Position == int_point) && (selected_DBPoint.IsErased == false))
                        {
                            ed.WriteMessage($"У полилинии на слое {polyline_Intersected.Layer} определяемая точка совпала с текущей точкой в чертеже\n");
                        }
                        else
                        {
                            all_points.Add(selected_DBPoint.Position);
                        }
                    }

                }
                else
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
            double S1 = MyCommonFunctions.Vychisli_LomDlinu_Poly(polyline_Intersected, new Point3d(point3D_before.X, point3D_before.Y, 0), new Point3d(int_point.X, int_point.Y, 0));
            // ed.WriteMessage($"Для анализа точки пересечения найдена предыдущая точка {point3D_before}\nПолучено расстояние по кривой от предыдущей точки {S1}\n");
            double S2 = MyCommonFunctions.Vychisli_LomDlinu_Poly(polyline_Intersected, new Point3d(point3D_after.X, point3D_after.Y, 0), new Point3d(int_point.X, int_point.Y, 0));
            //  ed.WriteMessage($"Для анализа точки пересечения найдена последующая точка {point3D_after}\nПолучено расстояние по кривой до последующей точки {S2}\n");
            double Hfull = point3D_after.Z - point3D_before.Z;
            double Spol = S1 + S2;
            double TanA = Hfull / Spol;
            double H1 = S1 * TanA;
            //  ed.WriteMessage($"Вычислена высота точки пересечения {point3D_before.Z + H1}\n");
            return point3D_before.Z + H1;
        }
        public Point3d find_Z(Point3d anyPoint3D, Point3dCollection sourceCol)
        {
            double myZ = double.MinValue;
            // Point3d point3D_onVert = new Point3d(0,0,myZ);
            foreach (Point3d col in sourceCol)
            {
                if (Round(anyPoint3D.X, 3) == Round(col.X, 3) && Round(anyPoint3D.Y, 3) == Round(col.Y, 3))
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
            double dX = point3D_2.X - point3D_1.X;
            double dY = point3D_2.Y - point3D_1.Y;
            double dZ = point3D_2.Z - point3D_1.Z;
            Point3d middle_point3D = new Point3d(point3D_1.X + 0.5 * dX, point3D_1.Y + 0.5 * dY, point3D_1.Z + 0.5 * dZ);
            return middle_point3D;
        }
        public double get_angle_90(Point3d point3D_1, Point3d point3D_2)
        {
            double angle_90 = 0;
            using (Line myLine = new Line())
            {
                myLine.StartPoint = point3D_1;
                myLine.EndPoint = point3D_2;
                angle_90 = myLine.Angle + 1.5708;
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
                    for (int i = 0; i < SortPointsGRcol.Count - 1; ++i)
                    {
                        Point3d point3D_1 = SortPointsGRcol[i];
                        Point3d point3D_2 = SortPointsGRcol[i + 1];
                        Point3d point_ins_MPtext = get_middle_point3D(point3D_1, point3D_2);//находим среднюю точку с помощью вспомогательной функции
                        ed.WriteMessage("Функция \"get_middle_point3D\" для середины междопутья отработала  успешно\n");
                        //для создания красивого текста междопутья осталось узнать угол поворота текста и его содержимое (расстояние между точками)
                        double text_MP_angle = get_angle_90(point3D_1, point3D_2);
                        ed.WriteMessage("Функция \"get_angle_90\" отработала успешно\n");
                        double text_MP_rasst = Vychisli_S(point3D_1, point3D_2);
                        string text_MP_string = text_MP_rasst.ToString("0.00", CultureInfo.InvariantCulture);//задаем содержимое правильного формата
                                                                                                             //все исходные данные известны, создаем текст в примерном положении и добавляем его в базу чертежа
                        TextStyleTable textStyleTableDoc = (TextStyleTable)Trans.GetObject(db.TextStyleTableId, OpenMode.ForRead, false, true);//открываем для чтения таблицу стилей текста
                        String sStyleName = textMP_style;

                        if (textStyleTableDoc.Has(sStyleName))
                        {
                            db.Textstyle = textStyleTableDoc[sStyleName];

                        }

                        DBText text_MP = new DBText();
                        acBlkTblRec.AppendEntity(text_MP);
                        Trans.AddNewlyCreatedDBObject(text_MP, true);
                        ed.WriteMessage($"Текст без содержимого создан  успешно\n");
                        //устанавливаем текущим стиль текста, переданный в функции, если он есть в базе чертежа

                        text_MP.TextString = text_MP_string;
                        text_MP.Rotation = text_MP_angle;
                        text_MP.Position = point_ins_MPtext;
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
        //________________________________________________________________________________________________________________
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
        public PromptEntityResult select_Entity(Type myType, String myMessage)
        {
            PromptEntityOptions promptPointoptions = new PromptEntityOptions(myMessage + "\n");
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

        public Point3d get_projection_on_rebro(Point3d anyPoint3D, Point3d point_1, Point3d point_2)
        {
            double myX = double.MinValue;
            double myY = double.MinValue;
            using (Line rebro = new Line(point_1, point_2))
            {
                Line b_line = new Line(point_1, anyPoint3D);
                double b = b_line.Length;
                double a = rebro.Length;
                Line c_line = new Line(point_2, anyPoint3D);
                double c = c_line.Length;
                double S = (a * a + b * b - c * c) / (2 * a);
                myX = point_1.X + ((S / a) * (point_2.X - point_1.X));
                myY = point_1.Y + ((S / a) * (point_2.Y - point_1.Y));
            }
            Point3d proj_point = new Point3d(myX, myY, anyPoint3D.Z);

            return proj_point;
        }
        public Point3d get_point_on_Face(Point3d anyPoint3D, Face anyFace)
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
            double H_on_Face = (H1 + H2 + H3) / 3;
            Point3d point_on_Face = new Point3d(anyPoint3D.X, anyPoint3D.Y, H_on_Face);
            return point_on_Face;
        }
        public Line point_on_rebro(Point3d anyPoint3D, Face anyFace)
        {
            Point3d vert_1 = new Point3d(anyFace.GetVertexAt(0).X, anyFace.GetVertexAt(0).Y, 0);
            Point3d vert_2 = new Point3d(anyFace.GetVertexAt(1).X, anyFace.GetVertexAt(1).Y, 0);
            Point3d vert_3 = new Point3d(anyFace.GetVertexAt(2).X, anyFace.GetVertexAt(2).Y, 0);
            Line line_1 = new Line(vert_1, vert_2);
            Line line_a = new Line(vert_1, new Point3d(anyPoint3D.X, anyPoint3D.Y, 0));
            Line line_b = new Line(vert_2, new Point3d(anyPoint3D.X, anyPoint3D.Y, 0));
            double my_Dist = line_a.Length + line_b.Length - line_1.Length;
            my_Dist = my_Dist / line_1.Length;
            if (my_Dist < 0.002)
            {
                Line line = new Line(anyFace.GetVertexAt(0), anyFace.GetVertexAt(1));
                line_1.Dispose();
                return line;
            }
            else
            {
                line_1 = new Line(vert_2, vert_3);
                line_a = new Line(vert_2, new Point3d(anyPoint3D.X, anyPoint3D.Y, 0));
                line_b = new Line(vert_3, new Point3d(anyPoint3D.X, anyPoint3D.Y, 0));
                my_Dist = line_a.Length + line_b.Length - line_1.Length;
                my_Dist = my_Dist / line_1.Length;
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
                    line_1 = new Line(vert_3, vert_1);
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
        public bool point_inside_face(Point3d anyPoint3D, Face anyFace)
        {
            bool inside = false;
            Point3d point_1 = new Point3d(anyFace.GetVertexAt(0).X, anyFace.GetVertexAt(0).Y, 0);
            Point3d point_2 = new Point3d(anyFace.GetVertexAt(1).X, anyFace.GetVertexAt(1).Y, 0);
            Point3d point_3 = new Point3d(anyFace.GetVertexAt(2).X, anyFace.GetVertexAt(2).Y, 0);

            if (point_1 != point_2 && point_2 != point_3)
            {
                Point3d med_point_1 = get_middle_point3D(point_2, point_3);
                Point3d med_point_2 = get_middle_point3D(point_1, point_3);
                Point3dCollection collection_mediana = new Point3dCollection();
                Line mediana_1 = new Line(point_1, med_point_1);
                Line mediana_2 = new Line(point_2, med_point_2);
                mediana_1.IntersectWith(mediana_2, Intersect.OnBothOperands, collection_mediana, IntPtr.Zero, IntPtr.Zero);
                if (collection_mediana.Count > 0)
                {
                    Point3d point_center_triangle = collection_mediana[0];
                    //ed.WriteMessage($"У грани получена центральная точка {point_center_triangle}\n");//----вспомогательно-----
                    Line Rasst_do_Point = new Line(point_center_triangle, new Point3d(anyPoint3D.X, anyPoint3D.Y, 0));
                    Point3dCollection collection = new Point3dCollection();
                    Line line_1 = new Line(new Point3d(point_1.X, point_1.Y, 0), new Point3d(point_2.X, point_2.Y, 0));
                    Line line_2 = new Line(new Point3d(point_3.X, point_3.Y, 0), new Point3d(point_2.X, point_2.Y, 0));
                    Line line_3 = new Line(new Point3d(point_1.X, point_1.Y, 0), new Point3d(point_3.X, point_3.Y, 0));
                    Rasst_do_Point.IntersectWith(line_1, Intersect.OnBothOperands, collection, IntPtr.Zero, IntPtr.Zero);
                    if ((collection.Count == 0) || (Round(collection[0].X, 3) == Round(anyPoint3D.X, 3) && Round(collection[0].Y, 3) == Round(anyPoint3D.Y, 3)))
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
            if (FaceSel != null)
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
                                Plane my_plane = anyFace.GetPlane();
                                Vector3d vector_Z = Vector3d.ZAxis;
                                Point3d point3D = new Point3d();
                                Line line_rebro = point_on_rebro(anyPoint3D, anyFace);
                                if (line_rebro != null)
                                {

                                    double dH = Vychisli_Z(line_rebro.StartPoint, anyPoint3D, line_rebro.EndPoint);
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
                                DBPoint new_Point = new DBPoint(point3D);
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

        public void SozdanieKodaTochkiPopera(BlockTableRecord acBlkTblRec, Point3d PopPointWithH, Entity popent, Line MyLine)
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
                    myBool = true;
                } //End If
                else
                {
                    myBool = false;
                }
            } //Next
            return myBool;
        } //End Function

     
        public Point3dCollection SortPoint3dCollection(Point3dCollection GotovyPoper3dCol)
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

        public static double Vychisli_S(Point3d p1, Point3d p2)
        {
            double S1 = 0;
            using (Line myLine = new Line(new Point3d(p1.X, p1.Y, 0), new Point3d(p2.X, p2.Y, 0)))
            {
                S1 = myLine.Length;
            }
            return S1;
        }// End Function
        public string SdelayKodTochkiPopera(Entity popent, Transaction Trans, bool TryVspomPoint)
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

        public double Vychisli_LomDlinu(Polyline3d myPolyline3d, Point3d myPoint, Point3d zeroPoint)
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




















