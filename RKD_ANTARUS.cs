using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.DatabaseServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using static FileHelper;
using static ЗавестиАнтарус.AutoScaleDrawingDimension;

namespace АнтарусРКД
{

    /// <summary>
    /// Класс создания РКД на насосные станции АНТАРУС
    /// </summary>
    public class AntarusRKD
    {

        #region Глобальные переменные для класса 
        private static ISldWorks swApp;
        private static ModelDoc2 swModel;
        private static DrawingDoc swDrawing;
        #endregion


        /// <summary>
        /// Основной метод создания РКД на насосную установку ANTARUS
        /// </summary>
        /// <param name="filepath"></param>
        public static void CreateRKD(string filepath)
        {

            Dictionary<string, string> paramANT = ReadParameters(filepath, '=');

            List<string> drawinglist=GetModelList(paramANT);

            foreach (string drawing in drawinglist)
            {
                CreatePDFDrawing(drawing, paramANT["Конфигурация"]);
            }
        }


        /// <summary>
        /// Создает PDF чертеж по пути до сборки/детали
        /// </summary>
        /// <param name="filepath">Путь до сборки/детали</param>
        /// <param name="config">Имя конфигурации</param>
        /// <returns></returns>
        private static bool CreatePDFDrawing(string filepath,string config)
        {
            swApp = SolidWorksManager.swApp;
            swModel = swApp.OpenDoc6(filepath, 2, 0, "", 0, 0);
            swModel.ShowConfiguration2(config);
            swModel.ForceRebuild3(false);
            double scaleDraw = GetScaleValue(swModel);
            string drawingpath=GetDrawingPath(filepath);
            swDrawing = (DrawingDoc)swApp.OpenDoc6(drawingpath, 3,0, "", 0, 0);
            AutoScaleDrawingA4(scaleDraw);
            return true;
        }

        /// <summary>
        /// Возвращает путь до чертежа, по пути до сборки/детали
        /// </summary>
        /// <param name="filepath">путь до детали/сборки</param>
        /// <returns></returns>
        private static string GetDrawingPath(string filepath) 
        {
            try
            {
                string drawingpath = Path.Combine(Path.GetFileNameWithoutExtension(filepath), ".SLDDRW");
                return drawingpath;
            }
            catch 
            {
                return null;
            } 
        }

        /// <summary>
        /// Возвращает список чертежей которые необходимы для РКД
        /// </summary>
        /// <param name="parameters"></param>
        /// <returns></returns>
        private static List<string> GetModelList(Dictionary<string, string> parameters) 
        {
            string material, pn;
            if (!parameters.TryGetValue("Материал", out material) ||
                !parameters.TryGetValue("PN", out pn))
                throw new ArgumentException("Нужны ключи \"Материал\" и \"PN\".");
            var key = material + "|" + pn;
            Func<List<string>> f;
            if (!Handlers.TryGetValue(key, out f))
                throw new NotSupportedException("Комбинация не поддержана: " + key);
            List<string> modellist = f();
            return modellist;
        }


        /// <summary>
        /// Словарь с наименованием файлов на которые необходимо выгрузить чертежи
        /// </summary>
        private static readonly Dictionary<string, Func<List<string>>> Handlers =
            new Dictionary<string, Func<List<string>>>(StringComparer.Ordinal)
            {
                { "ч.ст.|16", () =>
                    {
                        var files = new List<string>();
                        files.Add("");
                        files.Add("");
                        files.Add("");
                        return files;

                    }
                },

                { "ч.ст.|16/25", () =>
                    {
                        var files = new List<string>();
                        files.Add("");
                        files.Add("");
                        files.Add("");
                        return files;
                    }
                },

                { "ч.ст.|25", () =>
                    {
                        var files = new List<string>();
                        files.Add("");
                        files.Add("");
                        files.Add("");
                        return files;
                    }
                },

                { "н.ст.|16", () =>
                    {
                        var files = new List<string>();
                        files.Add("");
                        files.Add("");
                        files.Add("");
                        return files;
                    }

                },

                { "н.ст.|16/25", () =>
                    {
                        var files = new List<string>();
                        files.Add("");
                        files.Add("");
                        files.Add("");
                        return files;
                    }
                },

                { "н.ст.|25", () =>
                    {
                        var files = new List<string>();
                        files.Add("");
                        files.Add("");
                        files.Add("");
                        return files;
                    }
                }
            };

    }
}
