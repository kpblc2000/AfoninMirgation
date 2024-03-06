using System;
using System.Linq;
using Infrastructure;


#if NCAD
using HostMgd.ApplicationServices;
using HostMgd.EditorInput;
using Teigha.DatabaseServices;
using Teigha.Runtime;
using Teigha.Geometry;
#elif ACAD
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
#endif

namespace UsefulFunctionsNCad23.CadCommands
{
    public static class group_intersectionCmd
    {
        [CommandMethod("group_intersection", CommandFlags.UsePickSet)]
        public static void GroupIntersectionCommand()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

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
                            ObjectIdCollection int_pol_id = new ObjectIdCollection();
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
                                        if ((polyline_Intersected.Layer == "25" || polyline_Intersected.Layer == "Откосы" || polyline_Intersected.Layer == "12_Рельеф") && (polyline_Intersected.NumberOfVertices < 3)) continue;

                                        CommonMethods methods = new CommonMethods();
                                        methods.added_Vertex_Polyline(polyline_Cutting, polyline_Intersected);
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
                }
                else ed.WriteMessage("В стандартный набор не попало ни одной пересекаемой полилинии\n");

            }
            else ed.WriteMessage("Не выбрана секущая полилиния\n");

        }

    }
}
