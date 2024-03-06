
using System;
#if NCAD
using HostMgd.EditorInput;
using Teigha.DatabaseServices;
using Teigha.Geometry;
#elif ACAD
#endif

namespace UsefulFunctionsNCad23.Infrastructure
{
    internal class CommonMethods
    {
        public string SdelayOpisanieTochkiPopera_3(Entity popent, Transaction Trans)
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

        public String ProveritNalichiePodpisi(Entity popent)
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

        public double Vychisli_S(Point3d p1, Point3d p2)
        {
            double S1 = 0;
            using (Line myLine = new Line(new Point3d(p1.X, p1.Y, 0), new Point3d(p2.X, p2.Y, 0)))
            {
                S1 = myLine.Length;
            }
            return S1;
        }// End Function

        public void create_point_onFace(Point3d anyPoint3D, SelectionSet FaceSel)
        {
            BlockTable acBlkTbl;   //объявляем переменные для базы с примитивами чертежа 
            BlockTableRecord acBlkTblRec;
            if (FaceSel != null)
            {
                // ed.WriteMessage($"В работу попало {FaceSel.Count} граней\n");
                using (Transaction Trans = db.TransactionManager.StartTransaction())
                {
                    int count_made = 0;
                    acBlkTbl = (BlockTable)Trans.GetObject(db.BlockTableId, OpenMode.ForRead, false, true);      //открываем для чтения класс BlockTable
                    acBlkTblRec = Trans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false, true) as BlockTableRecord;
                    foreach (SelectedObject sObj in FaceSel)
                    {

                        Face anyFace = (Face)Trans.GetObject(sObj.ObjectId, OpenMode.ForRead, false, true);
                        if (anyFace != null)
                        {
                            bool check_point = point_inside_face(anyPoint3D, anyFace);
                            //  anyFace.UpgradeOpen();//-------вспомогательно--------
                            //anyFace.ColorIndex = 4;
                            //anyFace.LineWeight = LineWeight.LineWeight009;//-------вспомогательно--------
                            if (check_point)
                            {
                                // PlanarEntity my_plane = anyFace.GetPlane();// надо правильно находить плоскость
                                Plane my_plane = anyFace.GetPlane();
                                Vector3d vector_Z = Vector3d.ZAxis;
                                Point3d point3D = new Point3d();
                                Line line_rebro = point_on_rebro(anyPoint3D, anyFace);
                                if (line_rebro != null)
                                {

                                    double dH = Vychisli_Z(line_rebro.StartPoint, anyPoint3D, line_rebro.EndPoint);
                                    point3D = new Point3d(anyPoint3D.X, anyPoint3D.Y, line_rebro.StartPoint.Z + dH);
                                    if (point3D == null) ed.WriteMessage("Не сработал метод Vychisli_Z\n");
                                    line_rebro.Dispose();
                                    //acBlkTblRec.AppendEntity(line_rebro);
                                    //Trans.AddNewlyCreatedDBObject(line_rebro, true);
                                    //line_rebro.ColorIndex = 5;
                                    //line_rebro.LineWeight = LineWeight.LineWeight035;
                                    // ed.WriteMessage($"Точка создана на ребре {line_rebro.StartPoint} - {line_rebro.EndPoint}\n");
                                }
                                else
                                {
                                    point3D = anyPoint3D.Project(my_plane, vector_Z);
                                    if (point3D == null) ed.WriteMessage("Не сработал метод anyPoint3D.Project\n");
                                    // ed.WriteMessage($"Точка создана на грани, отметка {anyPoint3D.Z} \n");
                                    // anyFace.UpgradeOpen();
                                    //anyFace.ColorIndex = 1;
                                    //anyFace.LineWeight = LineWeight.LineWeight035;
                                }
                                DBPoint new_Point = new DBPoint(point3D);
                                ++count_made;
                                acBlkTblRec.AppendEntity(new_Point);
                                Trans.AddNewlyCreatedDBObject(new_Point, true);
                                Trans.Commit();
                                ed.WriteMessage($"Создана точка с координатами {new_Point.Position}\n");
                                my_plane.Dispose();
                                break;
                            }
                        }
                    }
                    // ed.WriteMessage($"Создано {count_made} точек\n");
                    if (count_made == 0)
                    {
                        ed.WriteMessage($"Точка {anyPoint3D} не попала ни в одно ребро\n");
                        Trans.Abort();
                    }

                }
            }
            else
            {
                ed.WriteMessage("В чертеже не обнаружено 3-д граней\n");
                return;
            }

        }
        
        public void added_Vertex_Polyline(Polyline polyline_Cutting, Polyline polyline_Intersected)
        {
            if (polyline_Cutting.StartPoint.Z != polyline_Intersected.StartPoint.Z)
            {
                ed.WriteMessage($"Секущая полилиния с пересекаемой полилинией с ObjectId {polyline_Intersected.ObjectId} находятся на разных уровнях\n");
            }
            //  Polyline modified_Polyline=new Polyline();
            int myIndex = 0;
            bool need_DBPoint = false;
            Form1 form1 = new Form1();

            //if (form1.checkBox2.Checked && form1.checkBox2.CheckState==CheckState.Checked)
            if (form1.checkBox2.Checked)
            {
                need_DBPoint = true;
            }
            Point3dCollection intersect_point3dCol = new Point3dCollection();
            polyline_Cutting.IntersectWith(polyline_Intersected, Intersect.OnBothOperands, intersect_point3dCol, IntPtr.Zero, IntPtr.Zero);
            if (intersect_point3dCol.Count > 0) // если пересечения нашлись - идем дальше, иначе сообщаем, что не нашли пересечений
            {
                // ed.WriteMessage($"С линией с ObjectId: {polyline_Intersected.ObjectId} найдено {intersect_point3dCol.Count} точек пересечения:\n");
                //пока создадим вспомогательное сообщение с количеством и координатами найденных точек
                foreach (Point3d int_point in intersect_point3dCol)
                {
                    // ed.WriteMessage($"\t{int_point}\n");
                    // сначала проверяем, не совпала ли точка пересечения с существующей вершиной полилинии
                    // - в этом случае ничего добавлять не нужно
                    bool isVertex = is_point_in_Current_Vertex(int_point, polyline_Intersected);
                    if (isVertex == false)
                    {

                        //в отдельной функции находим индекс, после которого надо вставить в пересекаемую линию вершину
                        myIndex = find_addvertex_index(polyline_Intersected, int_point);
                        polyline_Intersected.UpgradeOpen();
                        polyline_Intersected.AddVertexAt(myIndex, new Point2d(int_point.X, int_point.Y), 0, 0, 0);
                        // ed.WriteMessage($"В пересекаемую полилинию добавлена вершина перед вершиной {myIndex}\n");

                    }
                    else
                    {
                        //если точка пересечения совпала с текущей вершиной, то ничего не делаем, переходим к следующей точке в коллекции
                        ed.WriteMessage($"Точка пересечения {int_point} совпала с существующей вершиной пересекаемой полилинии\n");
                        // continue;
                    }
                    //----------------------------начальная задумка сделана. Далее (опционально, если это указано в управляющей форме,
                    //---------------------вставляем вершину и в секущую полилинию, и создаем 3-Д точки на месте пересечений с отметками,
                    //-------------------------------------------интерполированными по пересекаемым линиям-----------------------------------
                    bool isVertex_Cutting = is_point_in_Current_Vertex(int_point, polyline_Cutting);
                    if (isVertex_Cutting == false)
                    {
                        int cutting_Index = find_addvertex_index(polyline_Cutting, int_point);
                        polyline_Cutting.UpgradeOpen();
                        polyline_Cutting.AddVertexAt(cutting_Index, new Point2d(int_point.X, int_point.Y), 0, 0, 0);
                        //  ed.WriteMessage($"В секущую полилинию добавлена вершина перед вершиной {myIndex}\n");

                    }
                    else
                    {
                        ed.WriteMessage($"Точка пересечения {int_point} совпала с существующей вершиной секущей полилинии\n");
                        // continue;
                    }
                    //опционально создаем 3-д точки
                    //--------далее запускаем процедуру создания точки в полученном пересечении----------------------
                    //--для создания точки с отметкой у меня есть ее Х и Y (int_point),
                    //для вычисления отметки надо найти на нужной структурной линии (polyline_Intersected) ее такую вершину,
                    //у которой есть существующая DBPoint с отметкой (по линии дальше и по линии ближе
                    //к началу от найденной точки пересечения)--------------------------------------------------------
                    if (need_DBPoint)
                    {
                        // if (isVertex) myIndex= polyline_Intersected.GetPoint3dAt(int_point)
                        if (isVertex) myIndex = find_addvertex_index(polyline_Intersected, int_point);
                        createPointOnPolyline(int_point, polyline_Intersected, myIndex); // решить проблему получения индекса точки пересечения в существующей вершине
                    }
                    else ed.WriteMessage("Процедура создания точек пропущена из-за настроек\n");

                }

            }
            else
            {
                ed.WriteMessage($"Не нашлось пересечений с линией с ObjectId: {polyline_Intersected.ObjectId}\n");
                // return null;
            }
            //  return modified_Polyline;
        }

        public double Vychisli_LomDlinu(Polyline3d myPolyline3d, Point3d myPoint, Point3d zeroPoint)
        {
            double LomDlina = 0;
            Point3dCollection myPoint3DCollection = new Point3dCollection();//создаем вспомогательную коллекцию, куда будем
                                                                            //кидать каждую точку анализируемой 3-д полилинии
            try
            {
                using (Transaction Trans = db.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId acObjIdVert in myPolyline3d)                               //перебираем каждую вершину 3-д полилинии, это делается как ObjectId в 3-д полилинии
                    {
                        PolylineVertex3d Vert;   // объявляем переменную для каждой 3-д вершины
                        Vert = Trans.GetObject(acObjIdVert, OpenMode.ForRead) as PolylineVertex3d; //считываем вершину по её ObjectId
                        Point3d anyPoint3D = new Point3d(Vert.Position.X, Vert.Position.Y, Vert.Position.Z);
                        myPoint3DCollection.Add(anyPoint3D);
                    }//кидаем следующую вершину
                }

                Point3dCollection lomPoint3DCollection = new Point3dCollection();

                int stInd = 0, finInd = 0;
                if (myPoint3DCollection.IndexOf(myPoint) < myPoint3DCollection.IndexOf(zeroPoint))
                {
                    stInd = myPoint3DCollection.IndexOf(myPoint);
                    finInd = myPoint3DCollection.IndexOf(zeroPoint);
                }
                if (myPoint3DCollection.IndexOf(zeroPoint) < myPoint3DCollection.IndexOf(myPoint))
                {
                    stInd = myPoint3DCollection.IndexOf(zeroPoint);
                    finInd = myPoint3DCollection.IndexOf(myPoint);
                }

                //функция вычисляет длину ломаной линии на плоскости, ограниченную заданными вершинами исходной 3-d полилинии
                int j;
                for (j = stInd; j <= finInd; j++)
                {
                    Point3d addPoint = new Point3d(myPoint3DCollection[j].X, myPoint3DCollection[j].Y, 0);
                    lomPoint3DCollection.Add(addPoint);
                }
                using (Polyline3d lomLine = new Polyline3d(Poly3dType.SimplePoly, lomPoint3DCollection, false))
                {
                    LomDlina = lomLine.Length;
                }


            }
#if NCAD
            catch (Teigha.Runtime.Exception ex)
#else
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
#endif
            {
                ed.WriteMessage($"В процессе вычисления длины ломаной возникла ошибка: \n{ex}");

            }
            catch (System.Exception ex1)
            {
                ed.WriteMessage($"В процессе вычисления длины ломаной возникла ошибка: \n{ex1}");

            }
            return LomDlina;
        }//end function

        private bool point_inside_face(Point3d anyPoint3D, Face anyFace)
        {
            bool inside = false;
            Point3d point_1 = new Point3d(anyFace.GetVertexAt(0).X, anyFace.GetVertexAt(0).Y, 0);
            Point3d point_2 = new Point3d(anyFace.GetVertexAt(1).X, anyFace.GetVertexAt(1).Y, 0);
            Point3d point_3 = new Point3d(anyFace.GetVertexAt(2).X, anyFace.GetVertexAt(2).Y, 0);

            if (point_1 != point_2 && point_2 != point_3)
            {
                Point3d med_point_1 = get_middle_point3D(point_2, point_3);
                Point3d med_point_2 = get_middle_point3D(point_1, point_3);
                Point3dCollection collection_mediana = new Point3dCollection();
                Line mediana_1 = new Line(point_1, med_point_1);
                Line mediana_2 = new Line(point_2, med_point_2);
                mediana_1.IntersectWith(mediana_2, Intersect.OnBothOperands, collection_mediana, IntPtr.Zero, IntPtr.Zero);
                if (collection_mediana.Count > 0)
                {
                    Point3d point_center_triangle = collection_mediana[0];
                    //ed.WriteMessage($"У грани получена центральная точка {point_center_triangle}\n");//----вспомогательно-----
                    Line Rasst_do_Point = new Line(point_center_triangle, new Point3d(anyPoint3D.X, anyPoint3D.Y, 0));
                    Point3dCollection collection = new Point3dCollection();
                    Line line_1 = new Line(new Point3d(point_1.X, point_1.Y, 0), new Point3d(point_2.X, point_2.Y, 0));
                    Line line_2 = new Line(new Point3d(point_3.X, point_3.Y, 0), new Point3d(point_2.X, point_2.Y, 0));
                    Line line_3 = new Line(new Point3d(point_1.X, point_1.Y, 0), new Point3d(point_3.X, point_3.Y, 0));
                    Rasst_do_Point.IntersectWith(line_1, Intersect.OnBothOperands, collection, IntPtr.Zero, IntPtr.Zero);
                    if ((collection.Count == 0) || (Round(collection[0].X, 3) == Round(anyPoint3D.X, 3) && Round(collection[0].Y, 3) == Round(anyPoint3D.Y, 3)))
                    {
                        collection.Clear();
                        Rasst_do_Point.IntersectWith(line_2, Intersect.OnBothOperands, collection, IntPtr.Zero, IntPtr.Zero);
                        if ((collection.Count == 0) || (Round(collection[0].X, 3) == Round(anyPoint3D.X, 3) && Round(collection[0].Y, 3) == Round(anyPoint3D.Y, 3)))
                        {
                            collection.Clear();
                            Rasst_do_Point.IntersectWith(line_3, Intersect.OnBothOperands, collection, IntPtr.Zero, IntPtr.Zero);
                            if ((collection.Count == 0) || (Round(collection[0].X, 3) == Round(anyPoint3D.X, 3) && Round(collection[0].Y, 3) == Round(anyPoint3D.Y, 3)))
                            {
                                collection.Clear();
                                // ed.WriteMessage($"В работу ушла грань с центральной точкой {point_center_triangle}\n");
                                return true;
                            }
                            else
                            {
                                line_1.Dispose();
                                line_2.Dispose();
                                line_3.Dispose();
                                Rasst_do_Point.Dispose();
                                collection.Dispose();
                                return false;
                            }
                        }
                        else
                        {
                            line_1.Dispose();
                            line_2.Dispose();
                            line_3.Dispose();
                            Rasst_do_Point.Dispose();
                            collection.Dispose();
                            return false;
                        }
                    }
                    else
                    {
                        line_1.Dispose();
                        line_2.Dispose();
                        line_3.Dispose();
                        Rasst_do_Point.Dispose();
                        collection.Dispose();
                        // ed.WriteMessage($"Грань не подошла\n");
                        return false;
                    }


                }
                else
                {
                    ed.WriteMessage($"У грани с вершинами {point_1}\t{point_2}\t{point_3} не удалось найти медиану\n");//----вспомогательно-----
                }


            }

            return inside;
        }

        private Line point_on_rebro(Point3d anyPoint3D, Face anyFace)
        {
            Point3d vert_1 = new Point3d(anyFace.GetVertexAt(0).X, anyFace.GetVertexAt(0).Y, 0);
            Point3d vert_2 = new Point3d(anyFace.GetVertexAt(1).X, anyFace.GetVertexAt(1).Y, 0);
            Point3d vert_3 = new Point3d(anyFace.GetVertexAt(2).X, anyFace.GetVertexAt(2).Y, 0);
            Line line_1 = new Line(vert_1, vert_2);
            Line line_a = new Line(vert_1, new Point3d(anyPoint3D.X, anyPoint3D.Y, 0));
            Line line_b = new Line(vert_2, new Point3d(anyPoint3D.X, anyPoint3D.Y, 0));
            double my_Dist = line_a.Length + line_b.Length - line_1.Length;
            my_Dist = my_Dist / line_1.Length;
            if (my_Dist < 0.002)
            {
                Line line = new Line(anyFace.GetVertexAt(0), anyFace.GetVertexAt(1));
                line_1.Dispose();
                return line;
            }
            else
            {
                line_1 = new Line(vert_2, vert_3);
                line_a = new Line(vert_2, new Point3d(anyPoint3D.X, anyPoint3D.Y, 0));
                line_b = new Line(vert_3, new Point3d(anyPoint3D.X, anyPoint3D.Y, 0));
                my_Dist = line_a.Length + line_b.Length - line_1.Length;
                my_Dist = my_Dist / line_1.Length;
                if (my_Dist < 0.002)
                {
                    Line line = new Line(anyFace.GetVertexAt(1), anyFace.GetVertexAt(2));
                    line_1.Dispose();
                    line_a.Dispose();
                    line_b.Dispose();
                    return line;
                }
                else
                {
                    line_1 = new Line(vert_3, vert_1);
                    line_a = new Line(vert_3, new Point3d(anyPoint3D.X, anyPoint3D.Y, 0));
                    line_b = new Line(vert_1, new Point3d(anyPoint3D.X, anyPoint3D.Y, 0));
                    my_Dist = line_a.Length + line_b.Length - line_1.Length;
                    my_Dist = my_Dist / line_1.Length;
                    if (my_Dist < 0.002)
                    {
                        Line line = new Line(anyFace.GetVertexAt(2), anyFace.GetVertexAt(0));
                        line_1.Dispose();
                        line_a.Dispose();
                        line_b.Dispose();
                        return line;
                    }
                    else
                    {
                        line_1.Dispose();
                        line_a.Dispose();
                        line_b.Dispose();
                        // ed.WriteMessage($"Точка лежит внутри грани\n");
                        return null;
                    }
                }
            }
        }

        private void createPointOnPolyline(Point3d int_point, Polyline polyline_Intersected, int myIndex)
        {

            Point3dCollection all_points = new Point3dCollection();
            BlockTable acBlkTbl;   //объявляем переменные для базы с примитивами чертежа 
            BlockTableRecord acBlkTblRec;
            using (Transaction Trans = db.TransactionManager.StartTransaction()) // начинаем транзакцию
            {
                acBlkTbl = (BlockTable)Trans.GetObject(db.BlockTableId, OpenMode.ForRead, false, true);      //открываем для чтения класс BlockTable
                acBlkTblRec = Trans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false, true) as BlockTableRecord;//теперь мы получили доступ к пространству модели
                TypedValue[] TvPoints = new TypedValue[1];
                TvPoints.SetValue(new TypedValue((int)(DxfCode.Start), "POINT"), 0);
                SelectionFilter filterPoints = new SelectionFilter(TvPoints);
                PromptSelectionResult resultPoints = ed.SelectAll(filterPoints);
                //  PromptSelectionResult resultPoints = ed.SelectCrossingWindow(new Point3d(int_point.X + 100, int_point.Y - 100, 0), new Point3d(int_point.X - 100, int_point.Y + 100, 0), filterPoints);
                if (resultPoints.Status == PromptStatus.OK)
                {
                    foreach (SelectedObject Sobj in resultPoints.Value)// закидываем все точки чертежа в общую коллекцию
                    {
                        DBPoint selected_DBPoint = (DBPoint)Trans.GetObject(Sobj.ObjectId, OpenMode.ForRead, false);
                        if ((selected_DBPoint.Position == int_point) && (selected_DBPoint.IsErased == false))
                        {
                            ed.WriteMessage($"У полилинии на слое {polyline_Intersected.Layer} определяемая точка совпала с текущей точкой в чертеже\n");
                        }
                        else
                        {
                            all_points.Add(selected_DBPoint.Position);
                        }
                    }

                }
                else
                {
                    ed.WriteMessage("В текущем чертеже не найдено 3-d-точек\n");
                    Trans.Abort();
                }
                //после того, как заполнена коллекция (если в ней больше 1 объекта, идем по полилинии и ищем совпадение ее вершины с какой-нибудь точкой и если находим - берем ее отметку
                //  ed.WriteMessage($"В анализ берется {all_points.Count} точек\n");
                if (all_points.Count > 1)
                {
                    Point3d point3D_after = new Point3d();
                    Point3d point3D_before = new Point3d();
                    int count_made = 0;
                    for (int i = myIndex + 1; i < polyline_Intersected.NumberOfVertices; ++i)
                    {
                        point3D_after = find_Z(polyline_Intersected.GetPoint3dAt(i), all_points);
                        if (point3D_after.Z > double.MinValue) break;
                    }
                    if (point3D_after.Z > double.MinValue)
                    {
                        for (int j = myIndex - 1; j >= 0; --j)
                        {
                            point3D_before = find_Z(polyline_Intersected.GetPoint3dAt(j), all_points);
                            if (point3D_before.Z > double.MinValue) break;
                        }
                        if (point3D_before.Z > double.MinValue)
                        {
                            //после того, как нам стали известны 3-d точки на линии до и после пересечения, интерполяцией нужно найти отметку точки пересечения
                            double myZ = Vychisli_Z_on_Polyline(point3D_before, int_point, point3D_after, polyline_Intersected);
                            Point3d point3Dinterpol = new Point3d(int_point.X, int_point.Y, myZ);
                            DBPoint int_DPpoint = new DBPoint(point3Dinterpol);
                            acBlkTblRec.AppendEntity(int_DPpoint);
                            Trans.AddNewlyCreatedDBObject(int_DPpoint, true);
                            Trans.Commit();
                            ++count_made;
                            // return;
                            // break;
                        }
                        //else
                        //{
                        //    ed.WriteMessage($"У полилинии на слое {polyline_Intersected.Layer} не найдено подходящих точек до точки пересечения\n");
                        //  //  Trans.Abort();
                        //}


                        // break;  

                    }
                    //else
                    //{
                    //    ed.WriteMessage($"У полилинии на слое {polyline_Intersected.Layer} не найдено подходящих точек после точки пересечения\n");
                    //   // Trans.Abort();
                    //}

                    if (count_made == 0) ed.WriteMessage($"Для полилинии в слое \"{polyline_Intersected.Layer}\" не создано ни одной точки в месте пересечения\n");
                }
                else
                {
                    ed.WriteMessage("В текущем чертеже слишком мало 3-d-точек\n");
                    Trans.Abort();
                }
            }
        }

        private double Vychisli_Z_on_Polyline(Point3d point3D_before, Point3d int_point, Point3d point3D_after, Polyline polyline_Intersected)
        {
            double S1 = MyCommonFunctions.Vychisli_LomDlinu_Poly(polyline_Intersected, new Point3d(point3D_before.X, point3D_before.Y, 0), new Point3d(int_point.X, int_point.Y, 0));
            // ed.WriteMessage($"Для анализа точки пересечения найдена предыдущая точка {point3D_before}\nПолучено расстояние по кривой от предыдущей точки {S1}\n");
            double S2 = MyCommonFunctions.Vychisli_LomDlinu_Poly(polyline_Intersected, new Point3d(point3D_after.X, point3D_after.Y, 0), new Point3d(int_point.X, int_point.Y, 0));
            //  ed.WriteMessage($"Для анализа точки пересечения найдена последующая точка {point3D_after}\nПолучено расстояние по кривой до последующей точки {S2}\n");
            double Hfull = point3D_after.Z - point3D_before.Z;
            double Spol = S1 + S2;
            double TanA = Hfull / Spol;
            double H1 = S1 * TanA;
            //  ed.WriteMessage($"Вычислена высота точки пересечения {point3D_before.Z + H1}\n");
            return point3D_before.Z + H1;
        }

        private Point3d find_Z(Point3d anyPoint3D, Point3dCollection sourceCol)
        {
            double myZ = double.MinValue;
            // Point3d point3D_onVert = new Point3d(0,0,myZ);
            foreach (Point3d col in sourceCol)
            {
                if (Round(anyPoint3D.X, 3) == Round(col.X, 3) && Round(anyPoint3D.Y, 3) == Round(col.Y, 3))
                {
                    myZ = col.Z;
                    break;
                }
            }
            //ed.WriteMessage($"Получена точка с отметкой {myZ}\n");
            return new Point3d(anyPoint3D.X, anyPoint3D.Y, myZ);
        }

    }
}
