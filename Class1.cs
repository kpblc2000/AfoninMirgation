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
        //делаем функцию приведения всех характерных линий на отметку 0

       

        public bool is_point_in_Current_Vertex(Point3d point, Polyline polyline)
        {
            bool point_vertex = false;
            for (int i = 0; i < polyline.NumberOfVertices; ++i)
            {
                Point3d vertPoint = polyline.GetPoint3dAt(i);
                if (vertPoint == point)
                {
                    point_vertex = true;
                    break;
                }
            }
            return point_vertex;
        }

       public Point3d get_middle_point3D(Point3d point3D_1, Point3d point3D_2)
        {
            double dX = point3D_2.X - point3D_1.X;
            double dY = point3D_2.Y - point3D_1.Y;
            double dZ = point3D_2.Z - point3D_1.Z;
            Point3d middle_point3D = new Point3d(point3D_1.X + 0.5 * dX, point3D_1.Y + 0.5 * dY, point3D_1.Z + 0.5 * dZ);
            return middle_point3D;
        }
        public double get_angle_90(Point3d point3D_1, Point3d point3D_2)
        {
            double angle_90 = 0;
            using (Line myLine = new Line())
            {
                myLine.StartPoint = point3D_1;
                myLine.EndPoint = point3D_2;
                angle_90 = myLine.Angle + 1.5708;
            }
            return angle_90;
        }
        public void SdelayMPtexts(Point3dCollection SortPointsGRcol, string textMP_style, double textMP_height, string textMP_layer)
        {

            //метод создает тексты междопутий. Нужный слой уже создан в базе (перед выызовом процедуры).
            BlockTable acBlkTbl;   //объявляем переменные для базы с примитивами чертежа 
            BlockTableRecord acBlkTblRec;
            //  DocumentLock docklock = doc.LockDocument(DocumentLockMode.ExclusiveWrite,null,null,true);
            using (Transaction Trans = db.TransactionManager.StartTransaction()) // начинаем транзакцию
            {
                try //начинаем обработку с блоком конструкцией улавливания ошибок
                {
                    //Если есть указанный стиль в базе, то делать им.
                    acBlkTbl = (BlockTable)Trans.GetObject(db.BlockTableId, OpenMode.ForRead, false, true);      //открываем для чтения класс BlockTable
                    acBlkTblRec = Trans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false, true) as BlockTableRecord;//теперь мы получили доступ к пространству модели
                    //сначала получаем примерное положение текста посередине между точками коллекции
                    for (int i = 0; i < SortPointsGRcol.Count - 1; ++i)
                    {
                        Point3d point3D_1 = SortPointsGRcol[i];
                        Point3d point3D_2 = SortPointsGRcol[i + 1];
                        Point3d point_ins_MPtext = get_middle_point3D(point3D_1, point3D_2);//находим среднюю точку с помощью вспомогательной функции
                        ed.WriteMessage("Функция \"get_middle_point3D\" для середины междопутья отработала  успешно\n");
                        //для создания красивого текста междопутья осталось узнать угол поворота текста и его содержимое (расстояние между точками)
                        double text_MP_angle = get_angle_90(point3D_1, point3D_2);
                        ed.WriteMessage("Функция \"get_angle_90\" отработала успешно\n");
                        double text_MP_rasst = Vychisli_S(point3D_1, point3D_2);
                        string text_MP_string = text_MP_rasst.ToString("0.00", CultureInfo.InvariantCulture);//задаем содержимое правильного формата
                                                                                                             //все исходные данные известны, создаем текст в примерном положении и добавляем его в базу чертежа
                        TextStyleTable textStyleTableDoc = (TextStyleTable)Trans.GetObject(db.TextStyleTableId, OpenMode.ForRead, false, true);//открываем для чтения таблицу стилей текста
                        String sStyleName = textMP_style;

                        if (textStyleTableDoc.Has(sStyleName))
                        {
                            db.Textstyle = textStyleTableDoc[sStyleName];

                        }

                        DBText text_MP = new DBText();
                        acBlkTblRec.AppendEntity(text_MP);
                        Trans.AddNewlyCreatedDBObject(text_MP, true);
                        ed.WriteMessage($"Текст без содержимого создан  успешно\n");
                        //устанавливаем текущим стиль текста, переданный в функции, если он есть в базе чертежа

                        text_MP.TextString = text_MP_string;
                        text_MP.Rotation = text_MP_angle;
                        text_MP.Position = point_ins_MPtext;
                        text_MP.Layer = textMP_layer;
                        text_MP.Height = textMP_height;
                        text_MP.ColorIndex = 256;
                        text_MP.LineWeight = LineWeight.ByLayer;

                        //теперь перемещаем созданный текст, чтобы он был левее линии поперечника, и его середина была посередине линии
                        Vector3d myVector = text_MP.GeometricExtents.MinPoint.GetVectorTo(text_MP.Position);
                        text_MP.TransformBy(Matrix3d.Displacement(myVector));
                        Point3d middle_box = get_middle_point3D(text_MP.GeometricExtents.MaxPoint, text_MP.Position);
                        myVector = middle_box.GetVectorTo(text_MP.Position);
                        text_MP.TransformBy(Matrix3d.Displacement(myVector));

                        if (text_MP.TextStyleName == "ATP")
                        {
                            text_MP.Oblique = 0.2618;
                            text_MP.WidthFactor = 0.8;
                        }
                    }

                    Trans.Commit();

                }
#if NCAD
                catch (Teigha.Runtime.Exception ex)
#else
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
#endif
                {
                    ed.WriteMessage($"В процессе работы команды \"SdelayMPtexts\" обнаружена ошибка: {ex.Message}\n");
                    Trans.Abort();
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"В процессе работы команды \"SdelayMPtexts\" обнаружена ошибка: {ex.Message}\n");
                    Trans.Abort();
                }
            }

        }
        //________________________________________________________________________________________________________________
        public double Vychisli_Z(Point3d p1, Point3d p2, Point3d p3)
        {
            Line line1 = new Line(new Point3d(p1.X, p1.Y, 0), new Point3d(p2.X, p2.Y, 0));
            Line line2 = new Line(new Point3d(p2.X, p2.Y, 0), new Point3d(p3.X, p3.Y, 0));
            double S1 = line1.Length;
            double S2 = line2.Length;
            double Hfull = p3.Z - p1.Z;
            double Spol = S1 + S2;
            double TanA = Hfull / Spol;
            double H1 = S1 * TanA;
            line1.Dispose();
            line2.Dispose();
            return H1;
        }
        public PromptEntityResult select_Entity(Type myType, String myMessage)
        {
            PromptEntityOptions promptPointoptions = new PromptEntityOptions(myMessage + "\n");
            promptPointoptions.SetRejectMessage($"Выбран объект не {myType}!\n");
            promptPointoptions.AllowNone = false;
            promptPointoptions.AddAllowedClass(myType, true);
            PromptEntityResult promptResult_1 = ed.GetEntity(promptPointoptions);
            if (promptResult_1.Status == PromptStatus.OK)
            {
                return promptResult_1;
            }
            else
            {
                return null;
            }

        }
   
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




















