using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using static AutomatizationOfSW.AutoScaleDrawingDimension;
using static FileHelper;
using System.Runtime.InteropServices;
using System.Threading;

namespace ВыполнитьЗадачиSolidWorks
{
        /// <summary>
        /// Обнаружение новых файлов
        /// </summary>
    public class AntarusFileWatcher
    {
       
        private FileSystemWatcher watcher;

        /// <summary>
        /// Действие программы при запуске
        /// </summary>
        public void Start()
        {
            watcher = new FileSystemWatcher
            {
                Path = Costants.OnPerfFilePath,
                //Filter = "*.",
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite
            };
            watcher.Created += OnFileCreated;
            watcher.EnableRaisingEvents = true;
            ProcessExistingFiles();
            Log("Консольное приложение запущено, мониторинг папки: " + Costants.OnPerfFilePath);
            Console.WriteLine("Нажмите Enter для завершения работы.");
            Console.ReadLine();
        }

        /// <summary>
        /// Действия программы при обнаружении файла
        /// </summary>
        /// <param name="source"></param>
        /// <param name="e"></param>
        private void OnFileCreated(object source, FileSystemEventArgs e)
        {
            Log($"Обнаружен новый файл: {e.FullPath}");
            //Открытие SolidWorks
             
            //Запуск программы создания 3D и 2D моделей Антарус
            if (e.FullPath.Contains("ЗавестиАнтарус"))
            {
                RunWithTimeout(
                () => SolidWorksManager.swApp,  // Инициализация SolidWorks
                swApp => CreateDrawingAndModel(e.FullPath), // Вызов метода с параметрами
                "Не удалось создать 3D модель и(или) чертеж на: " + e.FullPath,
                90,// Таймаут в секундах
                e.FullPath
                );
            }

            //Запуск программы расчета нагрузок ББ
            else if (e.FullPath.Contains("РасчитатьНагрузки_ББ"))
            {
               RunWithTimeout(
               () => SolidWorksManager.swApp,  // Инициализация SolidWorks
               swApp => ProcessBBStaticLoad(e.FullPath, swApp), // Вызов метода с параметрами
               "Не удалось создать 3D модель и(или) чертеж на: " + e.FullPath,
               300,// Таймаут в секундах
                e.FullPath
               );
            }

            //Запуск программы чертежей фасадов
            else if (e.FullPath.Contains("ЧертежиФасадов_ББ"))
            {
                RunWithTimeout(
                () => SolidWorksManager.swApp,  // Инициализация SolidWorks
                swApp => ProcessFileDrawingFronts(e.FullPath, swApp), // Вызов метода с параметрами
                "Не удалось создать 3D модель и(или) чертеж на: " + e.FullPath,
                120,// Таймаут в секундах
                e.FullPath
                );
            }

            else
            {
                Console.WriteLine("Неизвестный тип файла: " + e.FullPath);
            }
        }

        /// <summary>
        /// Поиск имеющихся файлов в папке
        /// </summary>
        private void ProcessExistingFiles()
        {
            try
            {
                string[] files = Directory.GetFiles(Costants.OnPerfFilePath);

                foreach (var filePath in files)
                {
                    Log($"Обнаружен существующий файл: {Path.GetFileName(filePath)}");
                    OnFileCreated(this, new FileSystemEventArgs(WatcherChangeTypes.Created, Costants.OnPerfFilePath, Path.GetFileName(filePath)));
                }
            }
            catch (Exception ex)
            {
                Log("Ошибка при обработке существующих файлов: " + ex.Message);
            }
        }

        /// <summary>
        /// Выполнение программы получение чертежей фасадов на ББ
        /// </summary>
        /// <param name="filePath"></param>
        private void ProcessFileDrawingFronts(string filePath,ISldWorks swApp) 
        {

            //Копирование файла с параметрами в целевую папку
            CopyFileWithOverwrite(filePath, Costants.fileConfigBBFronts);
            //Запуск программы создания чертежей на фасады
            swApp.RunMacro(Costants.fileMacroBBFronts, "Фасады", "CreateDrawing");
        }
       
        /// <summary>
        /// Выполнение программы расчета запаса прочности на ББ
        /// </summary>
        /// <param name="filePath"></param>
        private void ProcessBBStaticLoad(string filePath, ISldWorks swApp)
        {
            //Копироние файла с параметрами в папку в целевую папку
            CopyFileWithOverwrite(filePath, Costants.fileConfigBBStaticLoad);
            //Запуск расчета прочности
            swApp.RunMacro(Costants.fileMacroBBStaticLoad, "НагрузкиББ", "SetSizes");
 
        }
        
        /// <summary>
        /// Запустить метод с таймаутом выполнения
        /// </summary>
        /// <param name="initSwApp">Экземпляр SolidWorks</param>
        /// <param name="action">Выполняемая задача</param>
        /// <param name="logMessage">Сообщение в случае неудачи</param>
        /// <param name="timeoutSeconds">Время на выполнение</param>
        /// <param name="e">Путь до найденного файла</param>
        public void RunWithTimeout(Func<ISldWorks> initSwApp, Action<ISldWorks> action, string logMessage, int timeoutSeconds, string e)
        {
            bool success = true;
            bool completed = false;
            int i = 1;
            Exception threadException = null;

            try
                {
                    while (completed == false && (i <= 2))
                    {
                    
                        Log($"Попытка {i} создать чертеж и модель");

                        ISldWorks swApp = initSwApp();

                        if (swApp == null)
                        {
                            Log("Ошибка инициализации SolidWorks.");
                            success = false;
                            break;
                        }
                        threadException = null;
                        Thread thread = new Thread(() =>
                        {
                            try
                            {
                                action(swApp);
                            }
                            catch (Exception ex)
                            {
                                threadException = ex;
                            }
                        });

                        thread.SetApartmentState(ApartmentState.STA);
                        thread.Start();

                        if (!thread.Join(TimeSpan.FromSeconds(timeoutSeconds)))
                        {
                            Log($"Попытка {i} неуспешна—{logMessage}");
                            success = false;
                            KillAllSolidWorksProcesses();
                        }
                        else
                        {
                            if (threadException != null)
                            {
                                Log($"Попытка {i} завершилась исключением: {threadException.Message}");
                                success = false;
                                swApp.CloseAllDocuments(true);
                            }
                            else 
                            {
                                Log($"Попытка {i} успешна!");
                                swApp.CloseAllDocuments(true);
                                completed = true;
                            }
                        }
                        i++;
                    }  
                }
                catch (Exception ex)
                {
                    Log($"Основное действие завершилось с ошибкой: {ex.Message}");
                    success = false;
                }

                finally
                {
                    //Формирование конечного пути файла с параметрами
                    string finalFilePath = Path.Combine(success ? Costants.filePathFinal : Costants.filePathFinalBad, Path.GetFileName(e));
                    if (finalFilePath.Contains("ЧертежиФасадов_ББ") || finalFilePath.Contains("РасчитатьНагрузки_ББ"))
                    {
                        finalFilePath = RemoveFileWithAddPostFix(finalFilePath);
                    }
                    //Удаление файла если он существует
                    if (File.Exists(finalFilePath))
                    {
                        File.Delete(finalFilePath);
                    }
                    //Перемещение файла
                    File.Move(e, finalFilePath);

                    //Записть в лог
                    Log($"Обработка файла завершена: {e}");
                }

        }

        public static void Main()
        {
                var app = new AntarusFileWatcher();
                app.Start();
        }
    }
}
