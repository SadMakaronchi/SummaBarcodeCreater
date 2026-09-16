namespace SummaBarcodeCreater
{
    public class Settings //параметры плагина
    {
            public string path_plt { get; set; } = @"C:\РЕЗКА";
            public int velosity { get; set; } = 600;
            public bool barc2 { get; set; } = true;
            public int overcut { get; set; } = 1;
            public bool smothing { get; set; } = true;
            public string[] color_name { get; set; } = new[] { "CutContour" };
            public bool doc_name { get; set; } = true;
            public bool plx {  get; set; } = true;
            public double up_bleed {  get; set; } = 25;
            public double down_bleed { get; set;} = 25;
            public double left_bleed { get;set; } = 10;
            public double right_bleed { get; set; } = 10;
        
    }
}
