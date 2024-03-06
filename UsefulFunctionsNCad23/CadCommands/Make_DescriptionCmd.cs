using System;
#if NCAD
using HostMgd.ApplicationServices;
using HostMgd.EditorInput;
using Teigha.Colors;
using Teigha.DatabaseServices;
using Teigha.Geometry;
using Teigha.Runtime;
#elif ACAD
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
#endif

namespace UsefulFunctionsNCad23.CadCommands
{
    public static class Make_DescriptionCmd
    {
        [CommandMethod("Make_Description", CommandFlags.Modal | CommandFlags.UsePickSet | CommandFlags.Session)]
        public static void Make_DescriptionCommand()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            BlockTable acBlkTbl;   //объявляем переменные для базы с примитивами чертежа 
            BlockTableRecord acBlkTblRec;
            using (Transaction Trans = db.TransactionManager.StartTransaction()) // начинаем транзакцию
            {
                DocumentLock docklock = doc.LockDocument(DocumentLockMode.AutoWrite, null, null, true);
                acBlkTbl = (BlockTable)Trans.GetObject(db.BlockTableId, OpenMode.ForRead, false, true);      //открываем для чтения класс BlockTable
                acBlkTblRec = Trans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false, true) as BlockTableRecord;
                TypedValue[] TvPoper = new TypedValue[2];
                TvPoper.SetValue(new TypedValue((int)(DxfCode.Start), "LWPOLYLINE,LINE,SPLINE"), 0);
                TvPoper.SetValue(new TypedValue((int)(DxfCode.LayerName), "1,2,3,4,5,6,7,8,10,12,13,14,20,23,24,28,29"), 1);
                SelectionFilter filterPoper = new SelectionFilter(TvPoper);
                PromptSelectionOptions promptSelectionOptions = new PromptSelectionOptions();
                promptSelectionOptions.MessageForAdding = "\nВыберите характерную линию\n";
                promptSelectionOptions.SingleOnly = true;
                PromptSelectionResult resultPoper = ed.GetSelection(promptSelectionOptions, filterPoper);
                if (resultPoper.Status == PromptStatus.OK)
                {
                    //_____________________с помощью транзакции добавляем в базу чертежа слой "подписи поперечников" (а если он там уже есть - то не добавляем)______________________________
                    LayerTable acLyrTbl = (LayerTable)Trans.GetObject(db.LayerTableId, OpenMode.ForRead);
                    String sLayerName = "подписи поперечников";

                    if (acLyrTbl.Has(sLayerName) == false)
                    {
                        LayerTableRecord acLyrTblRec = new LayerTableRecord();
                        // Устанавливаем слою нужный мне цвет по индексу 190 и ранее заданное имя
#if NCAD
                        acLyrTblRec.Color = Teigha.Colors.Color.FromColorIndex(ColorMethod.ByAci, 190);

#else
                        acLyrTblRec.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 190);
#endif
                        acLyrTblRec.Name = sLayerName;
                        // открываем таблицу слоев для записи
                        acLyrTbl.UpgradeOpen();
                        // записываем новый слой в таблицу слоев и в транзакцию
                        acLyrTbl.Add(acLyrTblRec);
                        Trans.AddNewlyCreatedDBObject(acLyrTblRec, true);
                    }
                    SelectionSet poperSel = resultPoper.Value;
                    ObjectId myObjectId = poperSel.GetObjectIds()[0];
                    Entity myEnt = (Entity)Trans.GetObject(myObjectId, OpenMode.ForWrite);
                    String myDescription = SdelayOpisanieTochkiPopera_3(myEnt, Trans);
                    //  ed.WriteMessage($"\nНашел описание линии:{myDescription}\n");
                    DBText opisText = new DBText();
                    Point3d opisTextInsPt = Find_FirstPoint(myEnt, Trans);
                    // ed.WriteMessage($"Координаты первой точки выбранной линии: {opisTextInsPt}\n");
                    bool ExistDescripInPoint = true;
                    ExistDescripInPoint = ProveritNalichiePodpisiInTochka(opisTextInsPt);
                    //  ed.WriteMessage($"Существует в начальной точке текущее описание: {ExistDescripInPoint}\n");
                    if (ExistDescripInPoint)
                    {
                        docklock.Dispose();
                        ObjectId idOpisText = Find_CurDescription(opisTextInsPt, Trans);
                        DBText old_opisText = (DBText)Trans.GetObject(idOpisText, OpenMode.ForWrite);
                        ed.WriteMessage($"Найдено старое описание {old_opisText.TextString}\n");
                        //opisText.TextString= myDescription;
                        old_opisText.Erase(true);
                        old_opisText.Dispose();


                    }

                    //------------------------------
                    opisText.TextString = myDescription;
                    opisText.Position = opisTextInsPt;
                    opisText.Layer = "подписи поперечников";
                    opisText.Justify = AttachmentPoint.BaseLeft;
                    opisText.Rotation = 0;
                    opisText.Height = 0.5;
                    opisText.ColorIndex = 190;
                    opisText.LineWeight = (LineWeight)0.3;
                    acBlkTblRec.AppendEntity(opisText);
                    Trans.AddNewlyCreatedDBObject(opisText, true);


                    //далее создаю диалог с выскакивающим окном для ввода желаемого текста описания.
                    //myDescription используется как текущий вариант по-умолчанию
                    //Form2 form2 = new Form2();

                    //Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(form2);
                    //form2.textBox1.Text = myDescription;
                    //DialogResult dialogOpis =form2.DialogResult;
                    //if (dialogOpis == DialogResult.Yes)
                    //{
                    //    myDescription = dialogOpis.ToString();
                    //    // myDescription = form2.returnString;
                    //}
                    //___________________________________________________________________
                    PromptStringOptions optionsOpis = new PromptStringOptions("Введите желаемое описание точки\n");
                    optionsOpis.DefaultValue = myDescription;
                    optionsOpis.UseDefaultValue = true;
                    optionsOpis.AllowSpaces = true;

                    PromptResult resultOpis = ed.GetString(optionsOpis);
                    if (resultOpis.Status != PromptStatus.OK)
                    {
                        ed.WriteMessage("Прервана работа команды\n");
                        return;
                    }
                    myDescription = resultOpis.StringResult;


                    opisText.TextString = myDescription;


                    //ed.WriteMessage($"\nБудем создавать описание {myDescription}\n");
                }
                else
                {
                    ed.WriteMessage("\nНе выбраны линии для создания подписи\n");
                }


                Trans.Commit();
                //docklock.Dispose();
            }



        }

    }
}
