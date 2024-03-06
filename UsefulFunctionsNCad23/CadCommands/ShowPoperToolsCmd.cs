#if NCAD
using Teigha.Runtime;
#elif ACAD
using Autodesk.AutoCAD.Runtime;
#endif

namespace UsefulFunctionsNCad23.CadCommands
{
    public static class ShowPoperToolsCmd
    {
        [CommandMethod("ShowPoperTools", CommandFlags.Modal)]

        public static void ShowPoperTools()
        {
            // Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(form1);
            Form1 form1 = new Form1();
#if NCAD
            HostMgd.ApplicationServices.Application.ShowModelessDialog(form1);
#else
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(form1);
#endif
            form1.Visible = true;

            //form1.Show();
            //form1.ShowDialog();
            //form1.ControlBox = true;

            //form1.Activate();
        }

    }
}
