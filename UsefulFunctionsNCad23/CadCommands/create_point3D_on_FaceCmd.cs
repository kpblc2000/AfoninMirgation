


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
    public static class create_point3D_on_FaceCmd
    {
        [CommandMethod("create_point3D_on_Face")]
        public static void create_point3D_on_Face()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            PromptPointOptions promptPointoptions = new PromptPointOptions("Укажите точку на экране, где хотите создать 3-д точку\n");
            promptPointoptions.AllowNone = false;
            PromptPointResult promptResult_1 = ed.GetPoint(promptPointoptions);
            if (promptResult_1.Status == PromptStatus.OK)
            {
                TypedValue[] TvFace = new TypedValue[1];
                TvFace.SetValue(new TypedValue((int)(DxfCode.Start), "3DFACE"), 0);
                SelectionFilter filterFace = new SelectionFilter(TvFace);
                // PromptSelectionResult resultFace = ed.SelectAll(filterFace);
                PromptSelectionResult resultFace = ed.SelectCrossingWindow(new Point3d(promptResult_1.Value.X + 100, promptResult_1.Value.Y - 100, 0), new Point3d(promptResult_1.Value.X - 100, promptResult_1.Value.Y + 100, 0), filterFace);
                SelectionSet FaceSel = resultFace.Value;

                CommonMethods methods = new CommonMethods();
                methods.create_point_onFace(promptResult_1.Value, FaceSel, db);
            }
        }

    }
}
