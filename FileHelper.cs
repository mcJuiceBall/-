﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using OfficeOpenXml;

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
    /// Копирование файла с перезаписыванием, если такой существует
    /// </summary>
    /// <param name="source"></param>
    /// <param name="destination"></param>
    public static void СopyAndDeleteSourceFile(string source, string destination)
    {
        if (File.Exists(destination))
        {
            File.Delete(destination);
        }
        File.Copy(source, destination);
        File.Delete(source);
    }


    public static string GetFileBMIParam(string folderpath) 
    {
        string[] files = Directory.GetFiles(folderpath, "*.txt");
        string fileName;
        Regex regex = new Regex(@"^[0-9]+\. ЧертежиФасадов_ББ\.txt$", RegexOptions.IgnoreCase);
        
        foreach (string file in files)
        {
            fileName = Path.GetFileName(file);

            if (regex.IsMatch(fileName))
            {
                return fileName;
            }
        }

        return null;
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
            renamed = Regex.Replace(renamed, " СоздатьЧертежБМИ", "");
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

        public static Dictionary<string, string> GetKeyValue(string filepath)
        {
            Dictionary<string,string> keyValuePairs = new Dictionary<string,string>();
            foreach (var key in File.ReadLines(filepath)) 
            {
                var part=key.Split('=');
                keyValuePairs[part[0].Trim()]=part[1].Trim();
            }
            return keyValuePairs;
        }

        public static object FindValueUnderMarker(ExcelWorksheet sheet, string marker)
        {
            int rows = sheet.Dimension.End.Row;
            int cols = sheet.Dimension.End.Column;

            
                for (int c = 1; c <= cols; c++)
                {
                    var cellVal = sheet.Cells[2, c].Text != null ? sheet.Cells[2, c].Text.Trim() : "";
                Console.WriteLine($"Ищем={marker}, Нашли={cellVal}");
                    if (!string.IsNullOrEmpty(cellVal) &&
                        cellVal.Equals(marker, StringComparison.OrdinalIgnoreCase))
                    {
                        // Берём ячейку под найденной
                        var targetCell = sheet.Cells[3, c];
                        return targetCell.Value;
                    }
                }

            return null;
        }



    /// <summary>
    /// Считывание параметров с текстового файла c разделителем 
    /// </summary>
    /// <param name="path">путь до файла с параметрами</param>
    /// <param name="splitter">Разделитель в файле</param>
    /// <returns>Словарь с ключами и параметрами</returns>
    public static Dictionary<string, string> ReadParameters(string path, char splitter)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            using (var reader = new StreamReader(path, Encoding.UTF8))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    int idx = line.IndexOf(splitter);
                    if (idx <= 0)
                        continue;

                    string key = line.Substring(0, idx).Trim();
                    string value = line.Substring(idx + 1).Trim();

                    result[key] = value;
                }
            }

            return result;
        }

}
