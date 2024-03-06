#if NCAD
using HostMgd.ApplicationServices;
using HostMgd.EditorInput;
using Application = HostMgd.ApplicationServices.Application;
#elif ACAD
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
#endif
using System;
using System.Runtime.CompilerServices;
using System.Windows.Forms;


namespace Infrastructure
{
    public class MessageService
    {
        public void ConsoleMessage(string Message, [CallerMemberName] string CallMethod = null)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                InfoMessage(Message, CallMethod);
                return;
            }

            Editor ed = doc.Editor;
            ed.WriteMessage("\n" + (string.IsNullOrWhiteSpace(CallMethod) ? "" : $"{CallMethod} : ") + Message);
        }

        public void InfoMessage(string Message, [CallerMemberName] string CallMethod = null)
        {
            MessageBox.Show(string.IsNullOrWhiteSpace(CallMethod) ? "" : $"{CallMethod} : " + Message, "Информация",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void ExceptionMessage(Exception ex, [CallerMemberName] string CallMethod = null)
        {
            MessageBox.Show("\n" + (string.IsNullOrWhiteSpace(CallMethod) ? "" : $"{CallMethod} : ") + ex.Message,
                "Системная ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        public void ExceptionMessage(string Message, Exception ex, [CallerMemberName] string CallMethod = null)
        {
            MessageBox.Show(
                "\n" + (string.IsNullOrWhiteSpace(CallMethod) ? "" : $"{CallMethod} : ") + Message + "\n" + ex.Message,
                "Системная ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
