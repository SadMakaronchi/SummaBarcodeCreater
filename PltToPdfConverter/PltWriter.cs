
using SummaBarcodeCreater;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

public static class PltWriter
{
    private const double HpglUnitsPerMm = 40.0; //константа единиц плоттера в 1мм(разрешщение резки)
    private const double MmPerInch = 25.4; //константа перевода из мм в дюймы
    private const double PointsPerInch = 72.0; //константа перевода из дюймов в пункты
    public static string Write(List<PdfContourExtractor.Contour> contours,double pageWidthPoints,double pageHeightPoints,bool rotate)
    {
        if (contours == null)
            throw new ArgumentNullException(nameof(contours));

        if (pageWidthPoints <= 0)
            throw new ArgumentOutOfRangeException(nameof(pageWidthPoints));

        var plt = new StringBuilder();
        foreach (var contour in contours)
        {
            if (contour == null || contour.Points.Count == 0)
                continue;

            // Переводим все точки контура в координаты HPGL.
            var points = new List<(int X, int Y)>();

            foreach (var point in contour.Points)
            {
                var converted = ToPlotterCoordinates(point.X,point.Y,pageWidthPoints,pageHeightPoints,rotate);

                // Не записываем последовательные одинаковые точки.
                if (points.Count == 0 || points[points.Count - 1] != converted)
                {
                    points.Add(converted);
                }
            }

            if (points.Count == 0)
                continue;

            // Перемещение к началу контура с поднятым ножом.
            plt.AppendLine($"PU{points[0].X} {points[0].Y};");

            // Для настоящей резки нужен как минимум один сегмент.
            if (points.Count > 1)
            {
                // Опускаем нож и проводим контур.
                for (int i = 1; i < points.Count; i++)
                {
                    plt.AppendLine($"PD{points[i].X} {points[i].Y};");
                }

                // Если контур замкнут, возвращаемся в начало,
                // только если последняя точка ещё не совпадает с первой.
                if (contour.IsClosed && points[points.Count - 1] != points[0])
                {
                    plt.AppendLine($"PD{points[0].X} {points[0].Y};");
                }
            }

            // Поднимаем нож перед следующим контуром.
            plt.AppendLine("PU;");
        }
        return plt.ToString();
    }

    private static (int X, int Y) ToPlotterCoordinates(double pdfX,
                                                       double pdfY,
                                                       double pageWidthPoints,
                                                       double pageHeightPoints,
                                                       bool rotate)
    {
        double scale = HpglUnitsPerMm * MmPerInch / PointsPerInch;

        if (rotate)
        {
            pdfX = pageWidthPoints - pdfX;
            pdfY = pageHeightPoints - pdfY;
        }

        // PDF: X вправо, Y вверх.
        // Плоттер: X вверх, Y влево.
        // Ноль плоттера находится справа внизу страницы.
        int x = (int)Math.Round(pdfY * scale);
        int y = (int)Math.Round((pageWidthPoints - pdfX) * scale);
        return (x, y);
    }
}
