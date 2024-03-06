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

        private static Point3d Find_FirstPoint(Entity myEntity, Transaction transaction)
        {
            //функция находит первую точку у полилинии, отрезка, сплайна
            Point3d first_point3D = new Point3d();

            if (myEntity is Polyline pline)
            {
                Point2d point = pline.GetPoint2dAt(0);
                return new Point3d(point.X, point.Y, 0.0);
            }

            if (myEntity is Line line)
            {
                return line.StartPoint;
            }

            throw new ArgumentOutOfRangeException();
        }

        private static string SdelayOpisanieTochkiPopera_3(Entity popent, Transaction Trans)
        {
            String OpisanieTochkiPopera = ProveritNalichiePodpisi(popent);
            //проверяем, не заданы ли заранее подписи. Если заданы - то должно сразу на Return,
            //если не заданы - то идет их создание "по-умолчанию"
            if (OpisanieTochkiPopera == "подпись не задана" || OpisanieTochkiPopera == "Подписей поперечников не существует")
            {
                switch (popent.GetType().ToString())
                {
#if NCAD
                    case "Teigha.DatabaseServices.Polyline":
#else
                    case "Autodesk.AutoCAD.DatabaseServices.Polyline":
#endif
                        using (Polyline ProverkaPolyLine = Trans.GetObject(popent.ObjectId, OpenMode.ForRead) as Polyline)
                        {
                            switch (ProverkaPolyLine.Linetype)
                            {
                                case "atp_122_v":
                                    OpisanieTochkiPopera = "водопровод";
                                    break;
                                case "atp_121_v":
                                    OpisanieTochkiPopera = "водопровод";
                                    break;
                                case "atp_122_d":
                                    OpisanieTochkiPopera = "дренаж";
                                    break;
                                case "atp_121_d":
                                    OpisanieTochkiPopera = "дренаж";
                                    break;
                                case "atp_122_k":
                                    OpisanieTochkiPopera = "канализация";
                                    break;
                                case "atp_121_k":
                                    OpisanieTochkiPopera = "канализация";
                                    break;
                                case "atp_122_kn":
                                    OpisanieTochkiPopera = "канализация напорная";
                                    break;
                                case "atp_122_kl":
                                    OpisanieTochkiPopera = "канализация ливневая";
                                    break;
                                case "atp_122_g":
                                    OpisanieTochkiPopera = "газопровод";
                                    break;
                                case "atp_121_g":
                                    OpisanieTochkiPopera = "газопровод";
                                    break;
                                case "atp_122_t":
                                    OpisanieTochkiPopera = "теплотрасса";
                                    break;
                                case "atp_121_t":
                                    OpisanieTochkiPopera = "теплотрасса";
                                    break;
                                case "atp_119_3":
                                    switch (ProverkaPolyLine.Layer)
                                    {
                                        case "7":
                                            OpisanieTochkiPopera = "кабель 0.4кВ";
                                            break;
                                        case "10":
                                            OpisanieTochkiPopera = "кабель СЦБ";
                                            break;
                                    } //End Select
                                    break;
                                case "atp_121_2":
                                    OpisanieTochkiPopera = "кабель 0.4кВ";
                                    break;
                                case "atp_120_3b":
                                    OpisanieTochkiPopera = "кабель 0.4кВ";
                                    break;
                                case "atp_119_1":
                                    OpisanieTochkiPopera = "кабель 10кВ";
                                    break;
                                case "atp_121_1":
                                    OpisanieTochkiPopera = "кабель 10кВ";
                                    break;
                                case "atp_120_3a":
                                    OpisanieTochkiPopera = "кабель 10кВ";
                                    break;
                                case "atp_133_t":
                                    OpisanieTochkiPopera = "кабель связи";
                                    break;
                                case "atp_133":
                                    OpisanieTochkiPopera = "кабель связи";
                                    break;
                                case "atp_121_3":
                                    OpisanieTochkiPopera = "кабель связи";
                                    break;
                                case "atp_122_vx":
                                    OpisanieTochkiPopera = "воздухопровод";
                                    break;
                                case "atp_121_vx":
                                    OpisanieTochkiPopera = "воздухопровод";
                                    break;
                                case "atp_5x2":
                                    OpisanieTochkiPopera = "край грунтовой дороги";
                                    break;
                                case "atp_472":
                                    OpisanieTochkiPopera = "железобетонный забор";
                                    break;
                                case "atp_473":
                                    OpisanieTochkiPopera = "железобетонный забор";
                                    break;
                                case "atp_474_2a":
                                    OpisanieTochkiPopera = "металлический забор";
                                    break;
                                case "atp_474_1b":
                                    OpisanieTochkiPopera = "металлический забор";
                                    break;
                                case "atp_475_1":
                                    OpisanieTochkiPopera = "деревянный забор";
                                    break;
                                case "atp_476_3":
                                    OpisanieTochkiPopera = "забор сетка-рабица";
                                    break;
                                case "atp_476_1":
                                    OpisanieTochkiPopera = "колючая проволока";
                                    break;
                                case "atp_477":
                                    OpisanieTochkiPopera = "изгородь";
                                    break;
                                case "atp_476_2a":
                                    OpisanieTochkiPopera = "гладкая проволока";
                                    break;
                                case "atp_475_4a":
                                    OpisanieTochkiPopera = "деревянный с опорами";
                                    break;
                                case "atp_475_2":
                                    OpisanieTochkiPopera = "деревянный решетчатый";
                                    break;
                                case "atp_278_6b":
                                    OpisanieTochkiPopera = "металлический парапет";
                                    break;
                                case "atp_278_5":
                                    OpisanieTochkiPopera = "каменный парапет";
                                    break;
                                case "atp_1.5x1.5":
                                    switch (ProverkaPolyLine.Layer)
                                    {
                                        case "20":
                                            OpisanieTochkiPopera = "край отмостки";
                                            break;
                                        case "23":
                                            OpisanieTochkiPopera = "край асфальта";
                                            break;
                                    }
                                    break;

                                case "Continuous":
                                    switch (ProverkaPolyLine.Layer)
                                    {
                                        case "20":
                                            OpisanieTochkiPopera = "стена здания";
                                            break;
                                        case "23":
                                            OpisanieTochkiPopera = "край дороги";
                                            break;
                                    }

                                    break;


                            }
                            switch (ProverkaPolyLine.Layer)
                            {
                                case "2":
                                    switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case 256:
                                            OpisanieTochkiPopera = "Хпр.ХкВ";
                                            break;
                                        case 1:
                                            OpisanieTochkiPopera = "Хпр.ХкВ";
                                            break;
                                    }
                                    break;

                                case "12":
                                    switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case int i when i != 7:
                                            OpisanieTochkiPopera = "смотри на плане";
                                            break;
                                    }
                                    break;
                                case "14":
                                    OpisanieTochkiPopera = "трубопровод спец.назначения";
                                    break;
                                case "23":
                                    switch (ProverkaPolyLine.Linetype)
                                    {
                                        case "atp_0x1":
                                            OpisanieTochkiPopera = "граница угодий";
                                            break;
                                        case "atp_0.2x1.5":
                                            OpisanieTochkiPopera = "граница угодий";
                                            break;
                                    }
                                    switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case 3:
                                            OpisanieTochkiPopera = "край канавы";
                                            break;
                                    }
                                    break;
                                case "28":
                                    switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case int i when i != 7:
                                            OpisanieTochkiPopera = "урез воды";
                                            break;
                                    }
                                    break;
                                case "24":
                                    switch (ProverkaPolyLine.Linetype)
                                    {
                                        case "atp_0x1":
                                            OpisanieTochkiPopera = "край деревьев";
                                            break;
                                        case "atp_0.2x1.5":
                                            OpisanieTochkiPopera = "край деревьев";
                                            break;
                                    }
                                    break;
                                case "29":
                                    switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case int i when i != 6:
                                            OpisanieTochkiPopera = "путь";
                                            break;
                                    }
                                    break;
                            }

                        } //end using

                        break;
#if NCAD
                    case "Teigha.DatabaseServices.Line":
#else

                    case "Autodesk.AutoCAD.DatabaseServices.Line":
#endif
                        using (Line ProverkaPolyLine = Trans.GetObject(popent.ObjectId, OpenMode.ForRead) as Line)
                        {
                            switch (ProverkaPolyLine.Linetype)
                            {
                                case "atp_122_v":
                                    OpisanieTochkiPopera = "водопровод";
                                    break;
                                case "atp_121_v":
                                    OpisanieTochkiPopera = "водопровод";
                                    break;
                                case "atp_122_d":
                                    OpisanieTochkiPopera = "дренаж";
                                    break;
                                case "atp_121_d":
                                    OpisanieTochkiPopera = "дренаж";
                                    break;
                                case "atp_122_k":
                                    OpisanieTochkiPopera = "канализация";
                                    break;
                                case "atp_121_k":
                                    OpisanieTochkiPopera = "канализация";
                                    break;
                                case "atp_122_kn":
                                    OpisanieTochkiPopera = "канализация напорная";
                                    break;
                                case "atp_122_kl":
                                    OpisanieTochkiPopera = "канализация ливневая";
                                    break;
                                case "atp_122_g":
                                    OpisanieTochkiPopera = "газопровод";
                                    break;
                                case "atp_121_g":
                                    OpisanieTochkiPopera = "газопровод";
                                    break;
                                case "atp_122_t":
                                    OpisanieTochkiPopera = "теплотрасса";
                                    break;
                                case "atp_121_t":
                                    OpisanieTochkiPopera = "теплотрасса";
                                    break;
                                case "atp_119_3":
                                    switch (ProverkaPolyLine.Layer)
                                    {
                                        case "7":
                                            OpisanieTochkiPopera = "кабель 0.4кВ";
                                            break;
                                        case "10":
                                            OpisanieTochkiPopera = "кабель СЦБ";
                                            break;
                                    } //End Select
                                    break;
                                case "atp_121_2":
                                    OpisanieTochkiPopera = "кабель 0.4кВ";
                                    break;
                                case "atp_120_3b":
                                    OpisanieTochkiPopera = "кабель 0.4кВ";
                                    break;
                                case "atp_119_1":
                                    OpisanieTochkiPopera = "кабель 10кВ";
                                    break;
                                case "atp_121_1":
                                    OpisanieTochkiPopera = "кабель 10кВ";
                                    break;
                                case "atp_120_3a":
                                    OpisanieTochkiPopera = "кабель 10кВ";
                                    break;
                                case "atp_133_t":
                                    OpisanieTochkiPopera = "кабель связи";
                                    break;
                                case "atp_133":
                                    OpisanieTochkiPopera = "кабель связи";
                                    break;
                                case "atp_121_3":
                                    OpisanieTochkiPopera = "кабель связи";
                                    break;
                                case "atp_122_vx":
                                    OpisanieTochkiPopera = "воздухопровод";
                                    break;
                                case "atp_121_vx":
                                    OpisanieTochkiPopera = "воздухопровод";
                                    break;
                                case "atp_5x2":
                                    OpisanieTochkiPopera = "край грунтовой дороги";
                                    break;
                                case "atp_472":
                                    OpisanieTochkiPopera = "железобетонный забор";
                                    break;
                                case "atp_473":
                                    OpisanieTochkiPopera = "железобетонный забор";
                                    break;
                                case "atp_474_2a":
                                    OpisanieTochkiPopera = "металлический забор";
                                    break;
                                case "atp_474_1b":
                                    OpisanieTochkiPopera = "металлический забор";
                                    break;
                                case "atp_475_1":
                                    OpisanieTochkiPopera = "деревянный забор";
                                    break;
                                case "atp_476_3":
                                    OpisanieTochkiPopera = "забор сетка-рабица";
                                    break;
                                case "atp_476_1":
                                    OpisanieTochkiPopera = "колючая проволока";
                                    break;
                                case "atp_477":
                                    OpisanieTochkiPopera = "изгородь";
                                    break;
                                case "atp_476_2a":
                                    OpisanieTochkiPopera = "гладкая проволока";
                                    break;
                                case "atp_475_4a":
                                    OpisanieTochkiPopera = "деревянный с опорами";
                                    break;
                                case "atp_475_2":
                                    OpisanieTochkiPopera = "деревянный решетчатый";
                                    break;
                                case "atp_278_6b":
                                    OpisanieTochkiPopera = "металлический парапет";
                                    break;
                                case "atp_278_5":
                                    OpisanieTochkiPopera = "каменный парапет";
                                    break;
                                case "atp_1.5x1.5":
                                    switch (ProverkaPolyLine.Layer)
                                    {
                                        case "20":
                                            OpisanieTochkiPopera = "край отмостки";
                                            break;
                                        case "23":
                                            OpisanieTochkiPopera = "край асфальта";
                                            break;
                                    }
                                    break;

                                case "Continuous":
                                    switch (ProverkaPolyLine.Layer)
                                    {
                                        case "20":
                                            OpisanieTochkiPopera = "стена здания";
                                            break;
                                        case "23":
                                            OpisanieTochkiPopera = "край дороги";
                                            break;
                                    }

                                    break;


                            }
                            switch (ProverkaPolyLine.Layer)
                            {
                                case "2":
                                    switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case 1:
                                            OpisanieTochkiPopera = "Хпр.ХкВ";
                                            break;
                                    }
                                    break;

                                case "12":
                                    switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case int i when i != 7:
                                            OpisanieTochkiPopera = "смотри на плане";
                                            break;
                                    }
                                    break;
                                case "14":
                                    OpisanieTochkiPopera = "трубопровод спец.назначения";
                                    break;
                                case "23":
                                    switch (ProverkaPolyLine.Linetype)
                                    {
                                        case "atp_0x1":
                                            OpisanieTochkiPopera = "граница угодий";
                                            break;
                                        case "atp_0.2x1.5":
                                            OpisanieTochkiPopera = "граница угодий";
                                            break;
                                    }
                                    switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case 3:
                                            OpisanieTochkiPopera = "край канавы";
                                            break;
                                    }
                                    break;
                                case "28":
                                    switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case int i when i != 7:
                                            OpisanieTochkiPopera = "урез воды";
                                            break;
                                    }
                                    break;
                                case "24":
                                    switch (ProverkaPolyLine.Linetype)
                                    {
                                        case "atp_0x1":
                                            OpisanieTochkiPopera = "край деревьев";
                                            break;
                                        case "atp_0.2x1.5":
                                            OpisanieTochkiPopera = "край деревьев";
                                            break;
                                    }
                                    break;
                                case "29":
                                    switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case int i when i != 6:
                                            OpisanieTochkiPopera = "путь";
                                            break;
                                    }
                                    break;
                            }

                        } //end using

                        break;
#if NCAD
                    case "Teigha.DatabaseServices.Spline":
#else
                    case "Autodesk.AutoCAD.DatabaseServices.Spline":
#endif
                        using (Spline ProverkaPolyLine = Trans.GetObject(popent.ObjectId, OpenMode.ForRead) as Spline)
                        {
                            switch (ProverkaPolyLine.Linetype)
                            {
                                case "atp_122_v":
                                    OpisanieTochkiPopera = "водопровод";
                                    break;
                                case "atp_121_v":
                                    OpisanieTochkiPopera = "водопровод";
                                    break;
                                case "atp_122_d":
                                    OpisanieTochkiPopera = "дренаж";
                                    break;
                                case "atp_121_d":
                                    OpisanieTochkiPopera = "дренаж";
                                    break;
                                case "atp_122_k":
                                    OpisanieTochkiPopera = "канализация";
                                    break;
                                case "atp_121_k":
                                    OpisanieTochkiPopera = "канализация";
                                    break;
                                case "atp_122_kn":
                                    OpisanieTochkiPopera = "канализация напорная";
                                    break;
                                case "atp_122_kl":
                                    OpisanieTochkiPopera = "канализация ливневая";
                                    break;
                                case "atp_122_g":
                                    OpisanieTochkiPopera = "газопровод";
                                    break;
                                case "atp_121_g":
                                    OpisanieTochkiPopera = "газопровод";
                                    break;
                                case "atp_122_t":
                                    OpisanieTochkiPopera = "теплотрасса";
                                    break;
                                case "atp_121_t":
                                    OpisanieTochkiPopera = "теплотрасса";
                                    break;
                                case "atp_119_3":
                                    switch (ProverkaPolyLine.Layer)
                                    {
                                        case "7":
                                            OpisanieTochkiPopera = "кабель 0.4кВ";
                                            break;
                                        case "10":
                                            OpisanieTochkiPopera = "кабель СЦБ";
                                            break;
                                    } //End Select
                                    break;
                                case "atp_121_2":
                                    OpisanieTochkiPopera = "кабель 0.4кВ";
                                    break;
                                case "atp_120_3b":
                                    OpisanieTochkiPopera = "кабель 0.4кВ";
                                    break;
                                case "atp_119_1":
                                    OpisanieTochkiPopera = "кабель 10кВ";
                                    break;
                                case "atp_121_1":
                                    OpisanieTochkiPopera = "кабель 10кВ";
                                    break;
                                case "atp_120_3a":
                                    OpisanieTochkiPopera = "кабель 10кВ";
                                    break;
                                case "atp_133_t":
                                    OpisanieTochkiPopera = "кабель связи";
                                    break;
                                case "atp_133":
                                    OpisanieTochkiPopera = "кабель связи";
                                    break;
                                case "atp_121_3":
                                    OpisanieTochkiPopera = "кабель связи";
                                    break;
                                case "atp_122_vx":
                                    OpisanieTochkiPopera = "воздухопровод";
                                    break;
                                case "atp_121_vx":
                                    OpisanieTochkiPopera = "воздухопровод";
                                    break;
                                case "atp_5x2":
                                    OpisanieTochkiPopera = "край грунтовой дороги";
                                    break;
                                case "atp_472":
                                    OpisanieTochkiPopera = "железобетонный забор";
                                    break;
                                case "atp_473":
                                    OpisanieTochkiPopera = "железобетонный забор";
                                    break;
                                case "atp_474_2a":
                                    OpisanieTochkiPopera = "металлический забор";
                                    break;
                                case "atp_474_1b":
                                    OpisanieTochkiPopera = "металлический забор";
                                    break;
                                case "atp_475_1":
                                    OpisanieTochkiPopera = "деревянный забор";
                                    break;
                                case "atp_476_3":
                                    OpisanieTochkiPopera = "забор сетка-рабица";
                                    break;
                                case "atp_476_1":
                                    OpisanieTochkiPopera = "колючая проволока";
                                    break;
                                case "atp_477":
                                    OpisanieTochkiPopera = "изгородь";
                                    break;
                                case "atp_476_2a":
                                    OpisanieTochkiPopera = "гладкая проволока";
                                    break;
                                case "atp_475_4a":
                                    OpisanieTochkiPopera = "деревянный с опорами";
                                    break;
                                case "atp_475_2":
                                    OpisanieTochkiPopera = "деревянный решетчатый";
                                    break;
                                case "atp_278_6b":
                                    OpisanieTochkiPopera = "металлический парапет";
                                    break;
                                case "atp_278_5":
                                    OpisanieTochkiPopera = "каменный парапет";
                                    break;
                                case "atp_1.5x1.5":
                                    switch (ProverkaPolyLine.Layer)
                                    {
                                        case "20":
                                            OpisanieTochkiPopera = "край отмостки";
                                            break;
                                        case "23":
                                            OpisanieTochkiPopera = "край асфальта";
                                            break;
                                    }
                                    break;

                                case "Continuous":
                                    switch (ProverkaPolyLine.Layer)
                                    {
                                        case "20":
                                            OpisanieTochkiPopera = "стена здания";
                                            break;
                                        case "23":
                                            OpisanieTochkiPopera = "край дороги";
                                            break;
                                    }

                                    break;


                            }
                            switch (ProverkaPolyLine.Layer)
                            {
                                case "2":
                                    switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case 1:
                                            OpisanieTochkiPopera = "Хпр.ХкВ";
                                            break;
                                    }
                                    break;

                                case "12":
                                    switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case int i when i != 7:
                                            OpisanieTochkiPopera = "смотри на плане";
                                            break;
                                    }
                                    break;
                                case "14":
                                    OpisanieTochkiPopera = "трубопровод спец.назначения";
                                    break;
                                case "23":
                                    switch (ProverkaPolyLine.Linetype)
                                    {
                                        case "atp_0x1":
                                            OpisanieTochkiPopera = "граница угодий";
                                            break;
                                        case "atp_0.2x1.5":
                                            OpisanieTochkiPopera = "граница угодий";
                                            break;
                                    }
                                    switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case 3:
                                            OpisanieTochkiPopera = "край канавы";
                                            break;
                                    }
                                    break;
                                case "28":
                                    switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case int i when i != 7:
                                            OpisanieTochkiPopera = "урез воды";
                                            break;
                                    }
                                    break;
                                case "24":
                                    switch (ProverkaPolyLine.Linetype)
                                    {
                                        case "atp_0x1":
                                            OpisanieTochkiPopera = "край деревьев";
                                            break;
                                        case "atp_0.2x1.5":
                                            OpisanieTochkiPopera = "край деревьев";
                                            break;
                                    }
                                    break;
                                case "29":
                                    switch (ProverkaPolyLine.ColorIndex)
                                    {
                                        case int i when i != 6:
                                            OpisanieTochkiPopera = "путь";
                                            break;
                                    }
                                    break;
                            }

                        } //end using

                        break;
                }



            }
            //в зависимости от того, с каким entity идет пересечение - создаем набор кодов по-умолчанию


            return OpisanieTochkiPopera;

        } //end function

        private static String ProveritNalichiePodpisi(Entity popent)
        {
            String PodpisString = "подпись не задана";
            //задаем слой, в котором будут располагаться подписи поперечников ("подписи поперечников").
            //Если в чертеже существует текст в данном слое, точка вставки которого эквивалентна начальной точке
            //пересекаемой полилинии, то содержимое текста передается в подпись.
            //Если такого текста нет, то передается условная строка "подпись не задана"
            TypedValue[] TvPodpis = new TypedValue[2];
            TvPodpis.SetValue(new TypedValue((int)(DxfCode.Start), "TEXT"), 0);
            TvPodpis.SetValue(new TypedValue((int)(DxfCode.LayerName), "подписи поперечников"), 1);
            SelectionFilter filterPodpis = new SelectionFilter(TvPodpis);
            PromptSelectionResult resultPodpis = ed.SelectAll(filterPodpis);
            SelectionSet PodpisSel = resultPodpis.Value;
            if (PodpisSel is null)
            {
                PodpisString = "Подписей поперечников не существует";
                return PodpisString;
                //Exit Function
            } //End If
            using (Transaction Trans = db.TransactionManager.StartTransaction())
            {
                try
                {
                    switch (popent.GetType().ToString())
                    {
#if NCAD
                        case "Teigha.DatabaseServices.Line":
#else
                        case "Autodesk.AutoCAD.DatabaseServices.Line":
#endif
                            using (Line ProverkaLine = Trans.GetObject(popent.ObjectId, OpenMode.ForRead) as Line)
                            {
                                foreach (SelectedObject PodpisObj in PodpisSel)
                                {
                                    DBText ProverkaText = Trans.GetObject(PodpisObj.ObjectId, OpenMode.ForRead) as DBText;
                                    if (ProverkaText.Position.X == ProverkaLine.StartPoint.X & ProverkaText.Position.Y == ProverkaLine.StartPoint.Y)
                                    {
                                        PodpisString = ProverkaText.TextString;
                                    } //End If
                                    ProverkaText.Dispose();
                                } //Next
                            } //end using
                            break;
#if NCAD
                        case "Teigha.DatabaseServices.Polyline":
#else
                        case "Autodesk.AutoCAD.DatabaseServices.Polyline":
#endif
                            using (Polyline ProverkaPolyLine = Trans.GetObject(popent.ObjectId, OpenMode.ForRead) as Polyline)
                            {
                                foreach (SelectedObject PodpisObj in PodpisSel)
                                {
                                    DBText ProverkaText = Trans.GetObject(PodpisObj.ObjectId, OpenMode.ForRead) as DBText;
                                    if (ProverkaText.Position.X == ProverkaPolyLine.StartPoint.X & ProverkaText.Position.Y == ProverkaPolyLine.StartPoint.Y)
                                    {
                                        PodpisString = ProverkaText.TextString;
                                    } //End If
                                    ProverkaText.Dispose();
                                } //Next
                            } //end using
                            break;
                    } //End Select
                }
#if NCAD
                catch (Teigha.Runtime.Exception ex)
#else
                catch (Autodesk.AutoCAD.Runtime.Exception ex )
#endif
                {
                    PodpisString = "В процессе поиска подписи поперечника произошла ошибка: " + ex.Message;
                    ed.WriteMessage(PodpisString);
                } //End Try
            }//end using
            return PodpisString;
        } //end function

        private static bool ProveritNalichiePodpisiInTochka(Point3d myPoint3d)
        {
            bool myBool = false;
            TypedValue[] TvOpisanie = new TypedValue[2];
            TvOpisanie.SetValue(new TypedValue((int)(DxfCode.Start), "TEXT"), 0);
            TvOpisanie.SetValue(new TypedValue((int)(DxfCode.LayerName), "подписи поперечников"), 1);
            SelectionFilter filterOpisanie = new SelectionFilter(TvOpisanie);
            PromptSelectionResult resultOpisanie = ed.SelectAll(filterOpisanie);
            if (resultOpisanie.Status == PromptStatus.OK)
            {
                SelectionSet poperOpisanie = resultOpisanie.Value;
                using (Transaction Trans = db.TransactionManager.StartTransaction())
                {
                    foreach (SelectedObject ObjOpisanie in poperOpisanie) //ошибка здесь
                    {
                        DBText TryOpisanie = Trans.GetObject(ObjOpisanie.ObjectId, OpenMode.ForRead, false, false) as DBText;
                        // ed.WriteMessage($"Сравниваются координаты текста описания: {TryOpisanie.Position}\n и точки {myPoint3d}\n");
                        if ((Round(TryOpisanie.Position.X, 2) == Round(myPoint3d.X, 2)) & (Round(TryOpisanie.Position.Y, 2) == Round(myPoint3d.Y, 2)))
                        {
                            myBool = true;
                            break;
                        } //End If
                        else
                        {
                            myBool = false;
                        }

                    }
                } //End Using
                // 
            }
            else ed.WriteMessage("При выполнении функции ProveritNalichiePodpisiInTochka не найдены подписи поперечников\n");

            return myBool;
        } //End Function

    }
}
