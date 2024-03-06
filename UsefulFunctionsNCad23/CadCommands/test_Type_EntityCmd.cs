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
    public static class test_Type_EntityCmd
    {
        [CommandMethod("test_Type_Entity", CommandFlags.UsePickSet)]
        public static void test_Type_Entity()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

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

    }
}
