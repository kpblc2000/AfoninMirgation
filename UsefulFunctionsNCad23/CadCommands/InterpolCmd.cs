﻿using System;
using Infrastructure;

#if NCAD
using HostMgd.ApplicationServices;
using HostMgd.EditorInput;
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
    public static class InterpolCmd
    {
        [CommandMethod("Interpol", CommandFlags.Modal | CommandFlags.UsePickSet | CommandFlags.Session)]
        public static void Interpol()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            BlockTable acBlkTbl;   //объявляем переменные для базы с примитивами чертежа 
            BlockTableRecord acBlkTblRec;
            Point3d point3D_1 = new Point3d();
            Point3d point3D_2 = new Point3d();
            Point3d point3D_3 = new Point3d();
            DocumentLock docklock = doc.LockDocument(DocumentLockMode.ExclusiveWrite, null, null, true);
            using (Transaction Trans = db.TransactionManager.StartTransaction()) // начинаем транзакцию
            {
                try
                {
                    acBlkTbl = (BlockTable)Trans.GetObject(db.BlockTableId, OpenMode.ForRead, false, true);      //открываем для чтения класс BlockTable
                    acBlkTblRec =
                        Trans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false, true) as
                            BlockTableRecord;

                    CommonMethods methods = new CommonMethods();

                    PromptEntityResult promptResult_1 = methods.select_Entity(typeof(DBPoint), "Выберите первую точку");
                    if (promptResult_1 != null)
                    {
                        DBPoint dBPoint_1 = (DBPoint)Trans.GetObject(promptResult_1.ObjectId, OpenMode.ForRead, false);
                        point3D_1 = dBPoint_1.Position;
                        ed.WriteMessage($"Получена первая точка {dBPoint_1.Position}\n");
                        PromptEntityResult promptResult_2 =
                            methods.select_Entity(typeof(DBPoint), "Выберите вторую точку");
                        if (promptResult_2 != null)
                        {
                            DBPoint dBPoint_2 =
                                (DBPoint)Trans.GetObject(promptResult_2.ObjectId, OpenMode.ForRead, false);
                            point3D_2 = dBPoint_2.Position;
                            ed.WriteMessage($"Получена вторая точка {dBPoint_2.Position}\n");
                            //Предлагаем выбрать полилинию, если пользователь откажется,
                            //то указываем вручную на экране, где создать 3-D точку
                            PromptEntityResult promptResult_3 = methods.select_Entity(typeof(Polyline),
                                "Выберите полилинию, на пересечении с которой построится точка");
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
                                    double inter_H = methods.Vychisli_Z(point3D_1, point3D_3, point3D_2);
                                    ed.WriteMessage($"Получено превышение {inter_H}\n");
                                    Point3d inter_Point3D = new Point3d(point3D_3.X, point3D_3.Y, point3D_1.Z + inter_H);
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
                                        double inter_H = methods.Vychisli_Z(point3D_1, point3D_3, point3D_2);
                                        ed.WriteMessage($"Получено превышение {inter_H}\n");
                                        Point3d inter_Point3D = new Point3d(point3D_3.X, point3D_3.Y, point3D_1.Z + inter_H);
                                        DBPoint newPoint = new DBPoint(inter_Point3D);
                                        acBlkTblRec.AppendEntity(newPoint);
                                        Trans.AddNewlyCreatedDBObject(newPoint, true);
                                        //теперь вставляем вершину в выбранную полилинию в этом месте
                                        int insert_place = find_addvertex_index(sec_polyline, inter_Point3D, methods);
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

        private static int find_addvertex_index(Polyline myPolyline, Point3d addedPoint, CommonMethods Methods)
        {
            //функция возвращает индекс вершины,
            //перед которой можно вставить в полилинию проверяемую точку, чтобы она не "перекрутилась" (по порядку)
            int numer = 0;
            for (int i = 0; i < myPolyline.NumberOfVertices - 1; ++i)
            {

                double S1 = Methods.Vychisli_S(myPolyline.GetPoint3dAt(i), myPolyline.GetPoint3dAt(i + 1));
                double S2 = Methods.Vychisli_S(myPolyline.GetPoint3dAt(i), addedPoint);
                if (S2 <= S1)
                {
                    numer = i + 1; break;
                }
            }
            return numer;
        }

    }
}
