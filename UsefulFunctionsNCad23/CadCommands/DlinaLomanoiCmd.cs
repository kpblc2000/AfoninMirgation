
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
    public static class DlinaLomanoiCmd

    {
        [CommandMethod("DlinaLomanoi", CommandFlags.Modal)]
        public static void DlinaLomanoi()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

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
    }
}
