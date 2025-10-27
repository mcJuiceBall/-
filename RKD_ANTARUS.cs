using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.DatabaseServices;
using SolidWorks.Interop.sldworks;
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
        #endregion


        /// <summary>
        /// Основной метод создания РКД на насосную установку ANTARUS
        /// </summary>
        /// <param name="filepath"></param>
        public static void CreateRKD(string filepath)
        {

            Dictionary<string, string> parametersANTARUS = ReadParameters(filepath, '=');
            
            swApp = SolidWorksManager.swApp;

            swModel = swApp.OpenDoc6(@"", 2, 0, "", 0, 0);

        }
    }
}
