using System;
using System.Globalization;
using System.Linq;
using static System.Math;



//using Autodesk.AutoCAD.Customization;
//using RibbonButton = Autodesk.AutoCAD.Customization.RibbonButton;
//using System.Activities.Expressions;

//[assembly: CommandClass(typeof(Useful_FunctionsCsh.AfoninCommands2))]


// This line is not mandatory, but improves loading performances

namespace Useful_FunctionsCsh
{
    public class Class1
    {
       //________________________________________________________________________________________________________________
        
    }
    // обработчик нажатия кнопки
    public class CommandHandler_ribbonButton1 : System.Windows.Input.ICommand
    {
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object param)
        {
            return true;
        }

        public void Execute(object parameter)
        {
#if NCAD
            HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("Sdelay_PoperCS" + " ", false, false, true);
#else
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("Sdelay_PoperCS" + " ", false, false, true);
#endif
        }
    }
    public class CommandHandler_ribbonButton2 : System.Windows.Input.ICommand
    {
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object param)
        {
            return true;
        }

        public void Execute(object parameter)
        {
#if NCAD
            HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("PrintPopFileCS" + " ", false, false, true);
#else
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("PrintPopFileCS" + " ", false, false, true);
#endif
        }
    }
    public class CommandHandler_ribbonButton3 : System.Windows.Input.ICommand
    {
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object param)
        {
            return true;
        }

        public void Execute(object parameter)
        {
#if NCAD
            HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("DlinaLomanoi" + " ", false, false, true);
#else
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("DlinaLomanoi" + " ", false, false, true);
#endif
        }
    }

    public class CommandHandler_ribbonButton4 : System.Windows.Input.ICommand
    {
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object param)
        {
            return true;
        }

        public void Execute(object parameter)
        {
#if NCAD
            HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("Test_Risui_Poper" + " ", false, false, true);
#else

            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("Test_Risui_Poper" + " ", false, false, true);
#endif
        }
    }

    public class CommandHandler_ribbonButton5 : System.Windows.Input.ICommand
    {
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object param)
        {
            return true;
        }

        public void Execute(object parameter)
        {
#if NCAD
            HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("ShowPoperTools" + " ", false, false, true);
#else
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("ShowPoperTools" + " ", false, false, true);
#endif
        }
    }

    public class CommandHandler_ribbonButton6 : System.Windows.Input.ICommand
    {
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object param)
        {
            return true;
        }

        public void Execute(object parameter)
        {
#if NCAD
            HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("Sdelay_Zero" + " ", false, false, true);
#else
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("Sdelay_Zero" + " ", false, false, true);
#endif
        }
    }

    public class CommandHandler_ribbonButton7 : System.Windows.Input.ICommand
    {
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object param)
        {
            return true;
        }

        public void Execute(object parameter)
        {
#if NCAD
            HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("Make_Description" + " ", false, false, true);
#else
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("Make_Description" + " ", false, false, true);
#endif
        }
    }

    public class CommandHandler_ribbonButton8 : System.Windows.Input.ICommand
    {
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object param)
        {
            return true;
        }

        public void Execute(object parameter)
        {
#if NCAD
            HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("Test_Intersection" + " ", false, false, true);
#else
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("Test_Intersection" + " ", false, false, true);
#endif
        }
    }
    public class CommandHandler_ribbonButton9 : System.Windows.Input.ICommand
    {
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object param)
        {
            return true;
        }

        public void Execute(object parameter)
        {
#if NCAD
            HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("Interpol" + " ", false, false, true);
#else
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("Interpol" + " ", false, false, true);
#endif
        }
    }
    public class CommandHandler_ribbonButton10 : System.Windows.Input.ICommand
    {
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object param)
        {
            return true;
        }

        public void Execute(object parameter)
        {
#if NCAD
            HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("group_intersection" + " ", false, false, true);
#else
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("group_intersection" + " ", false, false, true);
#endif
        }
    }
    public class CommandHandler_ribbonButton11 : System.Windows.Input.ICommand
    {
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object param)
        {
            return true;
        }

        public void Execute(object parameter)
        {
#if NCAD
            HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("create_point3D_on_each_vertex" + " ", false, false, true);
#else
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("create_point3D_on_each_vertex" + " ", false, false, true);
#endif
        }
    }
}




















