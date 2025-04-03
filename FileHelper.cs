using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;

    public class FileHelper
    {
        

        /// <summary>
    /// Ведедение лога действий
    /// </summary>
    /// <param name="message"></param>
        public static void Log(string message)
        {
            File.AppendAllText(Costants.logFile, $"{DateTime.Now}: {message}{System.Environment.NewLine}");
        }

    /// <summary>
    /// Копирование файла с перезаписыванием, если такой существует
    /// </summary>
    /// <param name="source"></param>
    /// <param name="destination"></param>
        public static void CopyFileWithOverwrite(string source, string destination)
        {
            if (File.Exists(destination))
            {
                File.Delete(destination);
            }
            File.Copy(source, destination);
        }

        /// <summary>
    /// Переименовать файл из формата ХХХХХХ.MLV ЗавестиАнтарус.xlsx в MLV.xlsx 
    /// </summary>
    /// <param name="filePath">Путь до файла</param>
    /// <returns></returns>
        public static string RenameFile(string filePath)
        {
            //Получить имя файла
            string fileName = Path.GetFileName(filePath);
            //Замена цифр в начале имени файла на ""
            string renamed = Regex.Replace(fileName, @"^\d+\.", "");
            //Убрать из названия "ЗавестиАнтарус"
            renamed = Regex.Replace(renamed, " ЗавестиАнтарус", "");
            //Новый путь файла
            string newPath = Path.Combine(Path.GetDirectoryName(filePath) ?? string.Empty, renamed);
            return newPath;
        }

        /// <summary>
        /// Перенос файла с добавлением даты в конце
        /// </summary>
        /// <param name="e"></param>
        /// <returns></returns>
        public static string RemoveFileWithAddPostFix(string filePath)
        {
            if (File.Exists(filePath))
            {
                DateTime dateTime = DateTime.Now;

                // Безопасный формат времени
                string safeTimestamp = dateTime.ToString("yyyy_MM_dd_HH_mm_ss");

                // Изменение имени файла
                string directory = Path.GetDirectoryName(filePath) ?? string.Empty;
                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);
                string extension = Path.GetExtension(filePath);

                // Новый путь с временным постфиксом
                string newFilePath = Path.Combine(directory, $"{fileNameWithoutExtension}+{safeTimestamp}{extension}");
                return newFilePath;
            }

            return filePath; // Если файл не существует, возвращаем исходный путь
        }

        /// <summary>
        /// Получить код установки в 1С из названия
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static string GetCodeFromFileName(string fileName)
        {
            Match match = Regex.Match(fileName, @"^\d+");
            return match.Success ? match.Value : string.Empty;
        }

}
