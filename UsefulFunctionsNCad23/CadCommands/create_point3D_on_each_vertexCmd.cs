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
    public static class create_point3D_on_each_vertexCmd
    {
        [CommandMethod("create_point3D_on_each_vertex", CommandFlags.Modal | CommandFlags.UsePickSet | CommandFlags.Session)]
        public static void create_point3D_on_each_vertex()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            BlockTable acBlkTbl;   //объявляем переменные для базы с примитивами чертежа 
            BlockTableRecord acBlkTblRec;
            PromptEntityResult result = select_Entity(typeof(Polyline), "Выберите полилинию, на вершинах которой надо создать точки\n");
            if (result != null)
            {
                using (Transaction Trans = db.TransactionManager.StartTransaction())
                {
                    DocumentLock docklock = doc.LockDocument();
                    acBlkTbl = (BlockTable)Trans.GetObject(db.BlockTableId, OpenMode.ForRead, false, true);      //открываем для чтения класс BlockTable
                    acBlkTblRec = Trans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false, true) as BlockTableRecord;
                    Polyline myPoly = (Polyline)Trans.GetObject(result.ObjectId, OpenMode.ForRead, false, true);
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

            }
            else
            {
                ed.WriteMessage("Не выбрана полилиния\n");
                return;
            }
        }


    }
}
