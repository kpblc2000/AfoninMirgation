
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
    public static class TestLomDlinaCmd
    {
        [CommandMethod("TestLomDlina", CommandFlags.Modal)]
        public static void TestLomDlina()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

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

                CommonMethods methods = new CommonMethods();

                double lomanaya = methods.Vychisli_LomDlinu(MyPl3d, myPoint, zeroPoint);

                MessageService msgService = new MessageService();
                msgService.ConsoleMessage($"Длина ломаной по полилинии между точками составила {lomanaya}");
                Trans.Commit();
            }
        }

    }
}
