using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DrawingConfigInterface
{
    public class DrawingConfig
    {
        public string[] NameOfView { get; set; }
        public string[] NotHideDimension { get; set; }

        public string[] DN_Collector { get; set; }
    }

    public static class DrawingConfigProvider
    {
        public static readonly Dictionary<string, DrawingConfig> Configurations = new Dictionary<string, DrawingConfig>
        {
            ["MLVФ.SLDDRW"] = new DrawingConfig
            {
                NameOfView = new[]
                {
                    "Чертежный вид11",
                    "Чертежный вид12",
                    "Чертежный вид13"
                },

                NotHideDimension = new[]
                {
                    "RD3@Чертежный вид13@MLVФ.Drawing",
                    "D2@Эскиз13@MLVФ.Drawing",
                    "RD1@Чертежный вид12@MLVФ.Drawing",
                    "RD2@Чертежный вид12@MLVФ.Drawing",
                    "RD3@Чертежный вид12@MLVФ.Drawing"
                }
            },

            ["MLVР.SLDDRW"] = new DrawingConfig
            {
                NameOfView = new[]
                {
                    "Чертежный вид1",
                    "Чертежный вид2",
                    "Чертежный вид3"
                },
                NotHideDimension = new[]
                {
                    "RD10@Чертежный вид3@MLVР.Drawing",
                    "RD1@Чертежный вид1@MLVР.Drawing",
                    "RD1@Чертежный вид2@MLVР.Drawing",
                    "RD3@Чертежный вид2@MLVР.Drawing",
                    "RD4@Чертежный вид2@MLVР.Drawing"
                }
            },

            ["MLV.SLDDRW"] = new DrawingConfig
            {
                NameOfView = new[]
                {
                    "Чертежный вид4",
                    "Чертежный вид6",
                    "Чертежный вид5"
                },
                NotHideDimension = new[]
                {
                    "RD4@Чертежный вид5@MLV.Drawing",
                    "RD1@Чертежный вид4@MLV.Drawing",
                    "RD1@Чертежный вид6@MLV.Drawing",
                    "RD3@Чертежный вид6@MLV.Drawing",
                    "RD4@Чертежный вид6@MLV.Drawing"
                }
            },

            ["IS.SLDDRW"] = new DrawingConfig
            {
                NameOfView = new[]
                {
                    "Чертежный вид12",
                    "Чертежный вид11",
                    "Чертежный вид15"
                },
                NotHideDimension = new[]
                {
                    "RD10@Чертежный вид15@IS.Drawing",
                    "RD1@Чертежный вид12@IS.Drawing",
                    "RD2@Чертежный вид12@IS.Drawing",
                    "RD1@Чертежный вид11@IS.Drawing",
                    "RD2@Чертежный вид11@IS.Drawing",
                    "RD7@Чертежный вид11@IS.Drawing"
                }
            },

            ["MST.SLDDRW"] = new DrawingConfig
            {
                NameOfView = new[]
                {
                    "Чертежный вид11",
                    "Чертежный вид12",
                    "Чертежный вид13"
                },
                NotHideDimension = new[]
                {
                    "RD4@Чертежный вид13@MST.Drawing",
                    "RD2@Чертежный вид11@MST.Drawing",
                    "RD3@Чертежный вид11@MST.Drawing",
                    "RD3@Чертежный вид12@MST.Drawing",
                    "RD4@Чертежный вид12@MST.Drawing",
                    "RD5@Чертежный вид12@MST.Drawing",
                    "RD2@Чертежный вид12@MST.Drawing",
                    "RD8@Чертежный вид12@MST.Drawing"
                }
            },

            ["MLHФ.SLDDRW"] = new DrawingConfig
            {
                NameOfView = new[]
                {
                    "Чертежный вид11",
                    "Чертежный вид12",
                    "Чертежный вид13"
                },
                NotHideDimension = new[]
                {
                    "RD4@Чертежный вид13@MLHФ.Drawing",
                    "D1@Эскиз13@MLHФ.Drawing",
                    "D2@Эскиз13@MLHФ.Drawing",
                    "RD3@Чертежный вид12@MLHФ.Drawing",
                    "RD10@Чертежный вид12@MLHФ.Drawing",
                    "RD11@Чертежный вид12@MLHФ.Drawing"
                }
            },

            ["MLHР.SLDDRW"] = new DrawingConfig
            {
                NameOfView = new[]
                {
                    "Чертежный вид11",
                    "Чертежный вид12",
                    "Чертежный вид13"
                },
                NotHideDimension = new[]
                {
                    "RD1@Чертежный вид13@MLHР.Drawing",
                    "D1@Эскиз13@MLHР.Drawing",
                    "D2@Эскиз13@MLHР.Drawing",
                    "RD1@Чертежный вид12@MLHР.Drawing",
                    "RD2@Чертежный вид12@MLHР.Drawing",
                    "RD3@Чертежный вид12@MLHР.Drawing",
                }
            },

            ["MLH.SLDDRW"] = new DrawingConfig
            {
                NameOfView = new[]
                {
                    "Чертежный вид12",
                    "Чертежный вид11",
                    "Чертежный вид13"
                },
                NotHideDimension = new[]
                {
                    "RD3@Чертежный вид13@MLH.Drawing",
                    "D1@Эскиз14@MLH.Drawing",
                    "D2@Эскиз14@MLH.Drawing",
                    "RD1@Чертежный вид11@MLH.Drawing",
                    "RD2@Чертежный вид11@MLH.Drawing",
                    "RD3@Чертежный вид11@MLH.Drawing"
                }
            }
        };
    }


    public static class DrawingConfigBMI
    {
        public static readonly Dictionary<string, DrawingConfig> Config = new Dictionary<string, DrawingConfig>
        {
            ["MLVФ.xlsx"] = new DrawingConfig
            {
                DN_Collector = new[]
                {
                    "$КОНФИГУРАЦИЯ@Фланец<6>",  //Всас
                    "$КОНФИГУРАЦИЯ@Фланец<5>",  //Напор
                    "$Состояние@Воротник<33>"    //Наличие жокея 
                }

            },

            ["MLVР.xlsx"] = new DrawingConfig
            {
                DN_Collector = new[]
                {
                    "$КОНФИГУРАЦИЯ@Фланец<54>", //Всас
                    "$КОНФИГУРАЦИЯ@Фланец<51>"  //Напор
                }
            },

            ["MLV.xlsx"] = new DrawingConfig
            {
                DN_Collector = new[]
                {
                    "$КОНФИГУРАЦИЯ@Фланец<6>", //Всас
                    "$КОНФИГУРАЦИЯ@Фланец<3>"  //Напор
                }
            },

            ["IS.xlsx"] = new DrawingConfig
            {
                DN_Collector = new[]
                {
                   "$КОНФИГУРАЦИЯ@Фланец<4>", //Всас
                   "$КОНФИГУРАЦИЯ@Фланец<11>" //Напор
                }
            },

            ["MST.xlsx"] = new DrawingConfig
            {
                DN_Collector = new[]
                {
                   "$КОНФИГУРАЦИЯ@Фланец<36>", //Всас
                   "$КОНФИГУРАЦИЯ@Фланец<41>" //Напор
                }
            },

            ["MLHФ.xlsx"] = new DrawingConfig
            {
                DN_Collector = new[]
                {
                   "$КОНФИГУРАЦИЯ@Фланец<4>", //Всас
                   "$КОНФИГУРАЦИЯ@Фланец<10>" //Напор
                }
            },

            ["MLHР.xlsx"] = new DrawingConfig
            {
                DN_Collector = new[]
                {
                   "$КОНФИГУРАЦИЯ@Фланец<45>", //Всас
                   "$КОНФИГУРАЦИЯ@Фланец<140>" //Напор
                }
            },

            ["MLH.xlsx"] = new DrawingConfig
            {
                DN_Collector = new[]
                {
                   "$КОНФИГУРАЦИЯ@Фланец<4>",   //Всас
                   "$КОНФИГУРАЦИЯ@Фланец<10>",  //Напор
                   "СтаринаЖокей"    //Наличие жоккея 
                }
            }
        };
    }
}
