#if NCAD
using HostMgd.ApplicationServices;
using HostMgd.EditorInput;
using Teigha.DatabaseServices;
using Teigha.Runtime;
#elif ACAD
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
#endif

namespace UsefulFunctionsNCad23.CadCommands
{
    public static class Test_IntersectionCmd
    {
        [CommandMethod("Test_Intersection", CommandFlags.UsePickSet)]
        public static void Test_IntersectionCommand()
        {

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

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
                                    if (polyline_Intersected != null)
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

                }
                else
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


    }
}
