using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using DrawingConfigInterface;
using static FileHelper;
using static System.Net.WebRequestMethods;
using System.Diagnostics;

namespace AutomatizationOfSW
{

    public class AutoScaleDrawingDimension
    {

        #region Глобальные переменные для класса 
        private static ISldWorks swApp;
        private static DrawingDoc swDrawing;
        private static ModelDoc2 swModel;
        private static string ModelPath;
        private static string DrawingPath;
        #endregion

        /// <summary>
        /// Завершает все процессы SolidWorks.
        /// </summary>
        public static void KillAllSolidWorksProcesses()
        {
            string[] SOLIDWORKS = new string[]
                                                {
                                                    "SLDWORKS",
                                                    "sldworks_fs",
                                                    "sldProcMon",
                                                    "swCefSubProc",
                                                 };

            foreach (var solid in SOLIDWORKS)
            {
                var solidWorksProcesses = Process.GetProcessesByName(solid);
                foreach (var proc in solidWorksProcesses)
                {
                    try
                    {
                        proc.Kill();
                        proc.WaitForExit(60000);
                        Log($"Процесс {proc.Id} успешно завершён.");
                    }
                    catch (Exception ex)
                    {
                        Log($"Ошибка при завершении процесса SolidWorks: {ex.Message}");
                    }
                }
            }
        }

        /// <summary>
        /// Попытка выполнить операцию с записью в лог результата
        /// </summary>
        /// <param name="operation"></param>
        /// <param name="operationName"></param>
        private static bool TryExecute(Action operation, string operationName)
        {
            try
            {
                operation();
                Log($"Попытка выполнить {operationName} успешна");
                return true;

            }
            catch(Exception ex)
            {
                Log($"Попытка  выполнить {operationName} не успешна.Ошибка:{ex.Message}");
                return false;
            }

        }

        /// <summary>
        /// Открывает Solidworks и возвращает приложение SOLIDWORKS
        /// </summary>
        public static class SolidWorksManager
        {
            private static ISldWorks SwApp;

            public static ISldWorks swApp
            {
                get
                {
                    // Если экземпляр уже существует, проверяем его на работоспособность
                    if (SwApp != null)
                    {
                        try
                        {
                            // Попытка обращения к свойству ActiveDoc
                            var doc = SwApp.ActiveDoc;
                        }
                        catch (Exception)
                        {
                            // Если объект не отвечает, сбрасываем его
                            SwApp = null;
                        }
                    }
                    // Если экземпляр отсутствует или был сброшен, создаем новый
                    if (SwApp == null)
                    {
                        SwApp = OpenSolidWorks();
                        SwApp.Visible = true;
                    }
                    return SwApp;
                }
            }

            private static ISldWorks OpenSolidWorks()
            {
                ISldWorks app = null;
                try
                {
                    // Попытка получить запущенный экземпляр SolidWorks
                    app = (ISldWorks)Marshal.GetActiveObject("SldWorks.Application");
                }
                catch (COMException)
                {
                    // Если SolidWorks не запущен, создаем новый экземпляр
                    Type swType = Type.GetTypeFromProgID("SldWorks.Application");
                    app = (ISldWorks)Activator.CreateInstance(swType);
                    // Если необходимо, можно добавить задержку для полной инициализации
                }
                return app;
            }
        }

        /// <summary>
        /// Сохранение файла в SAT формате версии 3.0 (Открывается в AUTOCAD и Revit)
        /// </summary>
        /// <param name="filePath"></param>
        public static void ExportToSat(ISldWorks swApp, string filePath, string PathToSaving)
        {
            if (swApp == null)
            {
                Console.WriteLine("Ошибка: SolidWorks не запущен!");
                return;
            }

            IModelDoc2 modelDoc = swModel;

            if (modelDoc == null)
            {
                Console.WriteLine("Ошибка: Откройте документ в SOLIDWORKS!");
                return;
            }

            // Обработка расширения файла
            string newFilePath = filePath.ToUpper().EndsWith(".SLDASM")
                ? filePath.Replace(".SLDASM", ".SAT")
                : filePath.Replace(".SLDPRT", ".SAT");
            string NewFile = $@"{PathToSaving}\3D {Path.GetFileName(PathToSaving)}.SAT";
            int errors = 0;
            int warnings = 0;
            // Сохранение в SAT
            ModelDocExtension modelDoc52 = modelDoc.Extension;
            bool boolstatus = swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swAcisOutputVersion, 4);
            boolstatus = modelDoc52.SaveAs3(
                NewFile,
                (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                null,
                0,
                ref errors,
                ref warnings
            );

            if (boolstatus)
                Console.WriteLine($"Файл успешно сохранен: {Path.GetFileName(PathToSaving)}.xlsx");
            else
                Console.WriteLine($"Ошибка сохранения: {errors}, предупреждения: {warnings}");

        }

        /// <summary>
        /// Получение пути сохранения файла
        /// </summary>
        /// <param name="swDrawing"></param>
        /// <returns></returns>
        public static string GetCustomProperties(string NameOfProperties)
        {
            string CustomProperties = swModel.GetCustomInfoValue("Установка", NameOfProperties);
            return CustomProperties;
        }

        /// <summary>
        /// Сохранение файла в PDF и DWG в целевую папку
        /// </summary>
        /// <param name="swDrawing">текущий чертеж</param>
        public static void SaveAsPDFandDWG(string Cod1C)
        {
                string FilePath = $@"{Costants.RootDirectory}\{Cod1C}\2D {Cod1C}";
                ((ModelDoc2)swDrawing).SaveAs(FilePath + ".PDF");
                ((ModelDoc2)swDrawing).SaveAs(FilePath + ".DWG");
        }

        /// <summary>
        /// Автоматическая группировка и размещение размеров на виде
        /// </summary>
        /// <param name="swView">Текущий вид</param>
        /// <param name="swModel">активная модель</param>
        public static void AutoPositionDimension(View swView)
        {
            Object[] vDispDim = (Object[])swView.GetDisplayDimensions();
            for (int j = 0; j < vDispDim.Length; j++)
            {
                DisplayDimension swDispDim = (DisplayDimension)vDispDim[j];
                Annotation swAnn = (Annotation)swDispDim.GetAnnotation();
                if ((!swAnn.IsDangling()) & (swAnn.Visible == (int)swAnnotationVisibilityState_e.swAnnotationVisible))
                {
                    swAnn.Select3(true, null);
                }
            }
            ModelDocExtension swModelDocExt = ((ModelDoc2)swDrawing).Extension;
            swModelDocExt.AlignDimensions((int)swAlignDimensionType_e.swAlignDimensionType_AutoArrange, 0.001);
            ((ModelDoc2)swDrawing).ClearSelection2(true);
        }

        /// <summary>
        /// Выполнение обработки размеров по видам
        /// </summary>
        /// <param name="swDrawing">Чертеж</param>
        /// <param name="NameOfView">Наименования видов</param>
        /// <param name="NotHideDimension">Размеры исключения</param>
        public static void ActionWithDimension(string[] NameOfView, string[] NotHideDimension)
        {
            int i = 0;
            for (i = 0; i < NameOfView.Length; i++)
            {
                ((ModelDoc2)swDrawing).Extension.SelectByID2(NameOfView[i], "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                View swView = ((ModelDoc2)swDrawing).SelectionManager.GetSelectedObject6(1, -1);
                if (i != NameOfView.Length - 1)
                {
                    HideDimensionView(swView, NotHideDimension);
                }
                else 
                { 
                    HideDimensionView(swView, NotHideDimension, specialView: true);
                }

                AutoPositionDimension(swView);
            }
        }

        /// <summary>
        /// Скрыть лишние габаритные размеры
        /// </summary>
        /// <param name="swView">Наименование вида</param>
        /// <param name="NotHideDimension">Размеры исключения</param>
        public static void HideDimensionView(View swView, string[] NotHideDimension, bool specialView = false)
        {
            string MaxValueName = "";
            double MaxValue = 0;
            Object[] vDispDim = (Object[])swView.GetDisplayDimensions();
            for (int j = 0; j < vDispDim.Length; j++)
            {
                DisplayDimension swDispDim = (DisplayDimension)vDispDim[j];
                Annotation swAnn = (Annotation)swDispDim.GetAnnotation();
                Dimension swDim = swDispDim.GetDimension2(0);

                if (!swAnn.IsDangling() && Array.IndexOf(NotHideDimension, swDim.FullName) == -1)
                {
                    swAnn.Visible = (int)swAnnotationVisibilityState_e.swAnnotationVisible;
                    double[] CurrentValue = swDim.GetValue3((int)swInConfigurationOpts_e.swThisConfiguration, null);
                    if (CurrentValue[0] > MaxValue)
                    {
                        MaxValue = CurrentValue[0];
                        MaxValueName = swDim.FullName;
                    }
                }
            }
            bool hideMaxDimension = false;
            if (specialView && NotHideDimension.Length > 0)
            {
                for (int j = 0; j < vDispDim.Length; j++)
                {
                    DisplayDimension swDispDim = (DisplayDimension)vDispDim[j];
                    Dimension swDim = swDispDim.GetDimension2(0);
                    if (swDim.FullName == NotHideDimension[0])
                    {
                        double[] value = swDim.GetValue3((int)swInConfigurationOpts_e.swThisConfiguration, null);
                        if (value[0] >= MaxValue)
                        {
                            hideMaxDimension = true;
                        }
                        break;
                    }
                }
            }

            for (int k = 0; k < vDispDim.Length; k++)
            {
                DisplayDimension swDispDim = (DisplayDimension)vDispDim[k];
                Annotation swAnn = (Annotation)swDispDim.GetAnnotation();
                Dimension swDim = swDispDim.GetDimension2(0);
                if ((swDim.FullName != MaxValueName || hideMaxDimension) && Array.IndexOf(NotHideDimension, swDim.FullName) == -1)
                {
                    swAnn.Visible = (int)swAnnotationVisibilityState_e.swAnnotationHidden;
                }
            }
        }

        /// <summary>
        /// Автоматический масштаб чертежа в зависимости от габаритов
        /// </summary>
        /// <returns>Возвращает чертеж</returns>
        public static void AutoScaleDrawing()
        {
            double ScaleValue = double.Parse(GetCustomProperties("Габарит"));
            Sheet swSheet = swDrawing.GetCurrentSheet();
            try
            {
                swSheet.SetScale(1, 5 * Math.Round(ScaleValue / 400), false, false);
            }
            catch
            {
                Console.WriteLine("Не удалось применить масштаб");
            }
        }

        /// <summary>
        /// Определение типа установки и запуск ActionWithDimension (обработку размеров на чертеже)
        /// </summary>
        /// <param name="swDrawing">Чертеж</param>
        public static void HideExcessDimension()
        {
            string fileName = Path.GetFileName(DrawingPath);
            if (DrawingConfigProvider.Configurations.TryGetValue(fileName, out DrawingConfig config))
            {
                string[] NameOfView = config.NameOfView;
                string[] NotHideDimension = config.NotHideDimension;
                ActionWithDimension(NameOfView, NotHideDimension);
            }
            else
            {
                Console.WriteLine("Выбранный файл не поддерживается");
            }
        }

        /// <summary>
        /// Создание 3D модели в SAT формате
        /// </summary>
        /// <param name="OpenPath"></param>
        /// <param name="OpenPathSAT"></param>
        /// <param name="FinalFolder"></param>
        public static void Create3Dmodel(string OpenPath,string OpenPathSAT,string FinalFolder) 
        {
                swApp = SolidWorksManager.swApp;
                swModel = swApp.OpenDoc6(OpenPath, 2, 0, "", 0, 0);
                ModelPath = Path.Combine(Path.GetDirectoryName(swModel.GetPathName()), Path.GetFileNameWithoutExtension(swModel.GetPathName()) + ".SLDASM");
                DrawingPath = Path.Combine(Path.GetDirectoryName(ModelPath), Path.GetFileNameWithoutExtension(ModelPath) + ".SLDDRW");
                swModel.ShowConfiguration("Установка");
                ModelDoc2 swModelSAT = swApp.OpenDoc6(OpenPathSAT, (int)swDocumentTypes_e.swDocASSEMBLY, 1, "", 0, 0);
                ExportToSat(swApp, OpenPathSAT, FinalFolder);
                swApp.CloseDoc(OpenPathSAT);
        }

        /// <summary>
        /// Создание 3D и 2D
        /// </summary>
        /// <param name="e"></param>
        public static void CreateDrawingAndModel(string e) 
        {
            Log($"Начинается обработка файла: {e}");
            #region Определение переменных
            string filePath = Path.Combine(Costants.RootDirectorySW, Path.GetFileName(e));
            string newFileName = RenameFile(Path.Combine(Costants.RootDirectorySW, Path.GetFileName(e)));
            string Cod1C = GetCodeFromFileName(Path.GetFileNameWithoutExtension(e));
            string OpenPath = $@"{Costants.RootDirectorySW}\{Path.GetFileNameWithoutExtension(newFileName)}.SLDASM";
            string DrawOpenPath = $@"{Costants.RootDirectorySW}\{Path.GetFileNameWithoutExtension(newFileName)}.SLDDRW";
            string OpenPathSAT = $@"{Costants.RootDirectorySW}\{Path.GetFileNameWithoutExtension(newFileName)}(SAT).SLDASM";
            string FinalFolder = $@"S:\.CADAutomation\Модели\{Cod1C}";
            #endregion
            
            CopyFileWithOverwrite(e, newFileName);

            if (!Directory.Exists(FinalFolder))
            {
                Directory.CreateDirectory(FinalFolder);
            }

            if (!TryExecute(() => Create3Dmodel(OpenPath, OpenPathSAT, FinalFolder), "Создать 3D модель"))
            {
                throw new Exception("Ошибка при создании 3D модели");
            }
            swDrawing = (DrawingDoc)swApp.OpenDoc6(DrawingPath, 3, 0, "", 0, 0);
            if (!TryExecute(() => AutoScaleDrawing(), "Автоматическая установка масштаба"))
            {
                throw new Exception("Ошибка при установке масштаба");
            }
            if(!TryExecute(() => HideExcessDimension(),"Скрыть лишние размеры"))
            {
                throw new Exception("Ошибка при скрытии лишних размеров");
            }
            if (!TryExecute(() => SaveAsPDFandDWG(Cod1C), "Сохранить в DWG и PDF"))
            {
                throw new Exception("Ошибка при сохранении в DWG и PDF");
            }

        }
    }
}
