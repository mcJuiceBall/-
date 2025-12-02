using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Authentication.ExtendedProtection;
using System.Text;
using System.Threading.Tasks;

public static class ChannelUtils
{
    private static Dictionary<string, Dictionary<string, int>> channel = new Dictionary<string, Dictionary<string, int>>
    {
        ["12П"] = new Dictionary<string, int>
        {
            ["Высота"] = 120,
            ["ШиринаПолки"] = 52
        },

        ["14П"] = new Dictionary<string, int>
        {
            ["Высота"] = 140,
            ["ШиринаПолки"] = 58
        },

        ["16П"] = new Dictionary<string, int>
        {
            ["Высота"] = 160,
            ["ШиринаПолки"] = 64
        }
    };


    /// <summary>
    /// Возвращает основные размеры швеллера (Высота, ШиринаПолки)
    /// </summary>
    /// <param name="shveller">Наименование швеллера</param>
    /// <param name="properties">Свойство которое необходимо вернуть</param>
    /// <returns></returns>
    public static int Channel(string shveller, string properties)
    {
        return channel[shveller][properties];
    }


    /// <summary>
    /// Проверяет находится ли величина в промежутке  +-2мм для швеллера каждой высоты из ГОСТ
    /// </summary>
    /// <param name="Lmin"></param>
    /// <returns></returns>
    public static bool IsLminInAnyChannelRange(double Lmin)
    {
        foreach (var shv in channel)
        {
            int h = shv.Value["Высота"];
            if (Lmin >= h - 2 && Lmin <= h + 2)
                return true;
        }

        return false;
    }
}

public class PumpFrame
{
    /// <summary>
    /// Находит наружный диаметр трубы по условному и материалу
    /// </summary>
    /// <param name="DN"></param>
    /// <param name="material"></param>
    /// <returns></returns>
    public static int OutD(double DN, string material) 
    {
        var pipes = new Dictionary<string, Dictionary<string, int>>
        {
            ["DN40"] = new Dictionary<string, int>
            {
                ["нерж"] = 48,
                ["ч.ст"] = 48
            },

            ["DN50"] = new Dictionary<string, int>
            {
                ["нерж"] = 60,
                ["ч.ст"] = 57
            },

            ["DN65"] = new Dictionary<string, int>
            {
                ["нерж"] = 76,
                ["ч.ст"] = 76
            },

            ["DN80"] = new Dictionary<string, int>
            {
                ["нерж"] = 89,
                ["ч.ст"] = 89
            },

            ["DN100"] = new Dictionary<string, int>
            {
                ["нерж"] = 104,
                ["ч.ст"] = 108
            },

            ["DN125"] = new Dictionary<string, int>
            {
                ["нерж"] = 129,
                ["ч.ст"] = 133
            },

            ["DN150"] = new Dictionary<string, int>
            {
                ["нерж"] = 154,
                ["ч.ст"] = 159
            },

            ["DN200"] = new Dictionary<string, int>
            {
                ["нерж"] = 204,
                ["ч.ст"] = 219
            },

            ["DN250"] = new Dictionary<string, int>
            {
                ["нерж"] = 254,
                ["ч.ст"] = 273
            },

            ["DN300"] = new Dictionary<string, int>
            {
                ["нерж"] = 304,
                ["ч.ст"] = 325
            },

            ["DN350"] = new Dictionary<string, int>
            {
                ["нерж"] = 356,
                ["ч.ст"] = 377
            },

            ["DN400"] = new Dictionary<string, int>
            {
                ["нерж"] = 406,
                ["ч.ст"] = 426
            },

            ["DN500"] = new Dictionary<string, int>
            {
                ["нерж"] = 508,
                ["ч.ст"] = 530
            }
        };

        int D = pipes[$"DN{DN}"][material];
        return D;
    }

    public static string GetConfigFrame(int DN, double Lmin) 
    {
        if (DN <= 300)
        {
            if (Lmin > 360) return "Швеллер Г-образный";

            if (Lmin > 200) return "Швеллер трапеция";

            if (Lmin > 20) return "Стандартная подставка";

            return "Невозможно сделать чертеж";
        }
       
        if (ChannelUtils.IsLminInAnyChannelRange(Lmin))
        {
            return "Швеллер трапеция";
        }

        if (Lmin > 162)
        {
            return "Швеллер Г-образный";
        }

        return "Невозможно сделать чертеж";
    }
   
    public static void CreateDrawingPumpFrame(string filepath)
    {
        var key_drawing=FileHelper.GetKeyValue(filepath);
        double Dbl(string key) => double.Parse(key_drawing[key]);
        double Lp = Dbl("Lp");
        double L1p = Dbl("L1p");
        double Bp = Dbl("Bp");
        double B1p = Dbl("B1p");
        double A = Dbl("A");
        double Np = Dbl("Np");
        double Dp = Dbl("Dp");
        double B11 = Dbl("B11");
        double B12 = Dbl("B12");
        double Hp = Dbl("Hp");
        double Hpp = Dbl("Hpp");
        double DN = Dbl("DN");
        string shveller = key_drawing["Швеллер"];
        string material = key_drawing["МатериалКоллектора"];
        int D = OutD(DN, material);
        double Lmin = Hp - (ChannelUtils.Channel(shveller, "Высота") + D / 2);
    }
}
 