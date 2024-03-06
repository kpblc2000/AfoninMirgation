using System;
using System.Runtime;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using static System.Data.DataTable;
using System.Windows.Input;
using System.ComponentModel;
using static System.Math;
#if NCAD
using Teigha.DatabaseServices;
using Teigha.Runtime;
using Teigha.Geometry;
using HostMgd.ApplicationServices;
using HostMgd.EditorInput;
using HostMgd.Windows;

using Platform = HostMgd;
using PlatformDb = Teigha;
#else
using Autodesk.Windows;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Internal;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices.Filters;
using Platform = Autodesk.AutoCAD;
    using PlatformDb = Autodesk.AutoCAD;
#endif

[assembly: ExtensionApplication(typeof(Useful_FunctionsCsh.MyPlugin))]

namespace Useful_FunctionsCsh
{
    internal class MyPlugin : IExtensionApplication
    {
#if NCAD
        Document doc = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        Editor ed = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
        Database db = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database;
#else
        Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
        Editor ed = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor;
        Database db = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
#endif
        public void Initialize()
        {
            try
            {
                ed.WriteMessage("Загружен плагин с дополнительными функциями");
                doc.SendStringToExecute("StartUFCsh" + " ", false, false, true);
            }
#if NCAD
            catch (Teigha.Runtime.Exception ex)
#else
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
#endif
            {
                ed.WriteMessage("При загрузке плагина обнаружена ошибка: " + ex.Message);
            }
            catch (System.Exception ex1)
            {
                ed.WriteMessage("При загрузке плагина обнаружена ошибка: " + ex1.Message);
            }

        }

        public void Terminate()
        {

        }
    }
}
