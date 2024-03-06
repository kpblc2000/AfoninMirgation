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
    public static class Sdelay_ZeroCmd
    {
        [CommandMethod("Sdelay_Zero", CommandFlags.Modal | CommandFlags.Session)]
        public static void Sdelay_Zero()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

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

    }
}
