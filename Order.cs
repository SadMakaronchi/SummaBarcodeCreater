using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace SummaBarcodeCreater
{
    public class Order
    {
        // ---------------------------------------------------------
        // Точка PLT
        // ---------------------------------------------------------
        private class PltPoint
        {
            public double X;
            public double Y;

            public PltPoint(double x, double y)
            {
                X = x;
                Y = y;
            }
        }

        // ---------------------------------------------------------
        // Отдельная траектория
        // ---------------------------------------------------------
        private class PltContour
        {
            public List<string> Commands =
                new List<string>();

            public double StartX;
            public double StartY;

            public double MinX;
            public double MaxX;
            public double MinY;
            public double MaxY;

            public double CenterX
            {
                get
                {
                    return (MinX + MaxX) / 2.0;
                }
            }

            public double CenterY
            {
                get
                {
                    return (MinY + MaxY) / 2.0;
                }
            }
        }

        // ---------------------------------------------------------
        // Основной метод
        // ---------------------------------------------------------
        public string Process(string plt)
        {
            List<PltContour> contours = Parse(plt);

            List<PltContour> ordered =
                SortLikeGoSign(contours);

            return BuildPlt(ordered);
        }

        private string BuildPlt(List<PltContour> contours)
        {
            List<string> lines = new List<string>();

            foreach (PltContour contour in contours)
            {
                foreach (string command in contour.Commands)
                {
                    lines.Add(command);
                }
            }

            return string.Join(
                Environment.NewLine,
                lines);
        }

        // ---------------------------------------------------------
        // Разбор PLT
        // ---------------------------------------------------------
        private List<PltContour> Parse(string text)
        {
            List<PltContour> result =
                new List<PltContour>();

            PltContour current = null;

            string[] lines = text.Split(
                new[] { '\r', '\n' },
                StringSplitOptions.RemoveEmptyEntries);

            foreach (string rawLine in lines)
            {
                string line = rawLine.Trim();

                if (line.Length == 0)
                    continue;

                Match match = Regex.Match(
                    line,
                    @"^(PU|PD)\s*(-?\d+(?:[.,]\d+)?)[,\s]+(-?\d+(?:[.,]\d+)?)",
                    RegexOptions.IgnoreCase);

                if (!match.Success)
                    continue;

                string command =
                    match.Groups[1].Value.ToUpper();

                string sx = match.Groups[2].Value
                    .Replace(',', '.');

                string sy = match.Groups[3].Value
                    .Replace(',', '.');

                double x = double.Parse(
                    sx,
                    CultureInfo.InvariantCulture);

                double y = double.Parse(
                    sy,
                    CultureInfo.InvariantCulture);

                // -------------------------------------------------
                // PU начинает новую траекторию
                // -------------------------------------------------
                if (command == "PU")
                {
                    if (current != null)
                    {
                        result.Add(current);
                    }

                    current = new PltContour();

                    current.StartX = x;
                    current.StartY = y;

                    current.MinX = x;
                    current.MaxX = x;

                    current.MinY = y;
                    current.MaxY = y;

                    current.Commands.Add(line);
                }

                // -------------------------------------------------
                // PD продолжается в текущей траектории
                // -------------------------------------------------
                else if (command == "PD")
                {
                    if (current == null)
                        continue;

                    current.Commands.Add(line);

                    if (x < current.MinX)
                        current.MinX = x;

                    if (x > current.MaxX)
                        current.MaxX = x;

                    if (y < current.MinY)
                        current.MinY = y;

                    if (y > current.MaxY)
                        current.MaxY = y;
                }
            }

            if (current != null)
            {
                result.Add(current);
            }

            return result;
        }

        // ---------------------------------------------------------
        // Формирование порядка GoSign
        // ---------------------------------------------------------
        private List<PltContour> SortLikeGoSign(
            List<PltContour> contours)
        {
            List<PltContour> result =
                new List<PltContour>();

            if (contours.Count == 0)
                return result;

            // -----------------------------------------------------
            // Определяем расстояние между колонками.
            //
            // У нас контуры одного столбца имеют почти одинаковый X.
            // -----------------------------------------------------
            double xStep = EstimateXStep(contours);

            // Допуск для объединения в одну колонку.
            double xTolerance = xStep * 0.35;

            List<List<PltContour>> columns =
                new List<List<PltContour>>();

            // -----------------------------------------------------
            // Раскладываем контуры по колонкам
            // -----------------------------------------------------
            foreach (PltContour contour in contours)
            {
                List<PltContour> foundColumn = null;

                foreach (List<PltContour> column in columns)
                {
                    if (column.Count == 0)
                        continue;

                    double columnX =
                        column[0].CenterX;

                    if (Math.Abs(
                        columnX - contour.CenterX)
                        <= xTolerance)
                    {
                        foundColumn = column;
                        break;
                    }
                }

                if (foundColumn == null)
                {
                    foundColumn =
                        new List<PltContour>();

                    columns.Add(foundColumn);
                }

                foundColumn.Add(contour);
            }

            // -----------------------------------------------------
            // Колонки слева направо
            // -----------------------------------------------------
            columns.Sort(
                delegate (
                    List<PltContour> a,
                    List<PltContour> b)
                {
                    return a[0].CenterX
                        .CompareTo(b[0].CenterX);
                });

            // -----------------------------------------------------
            // Змейка GoSign
            // -----------------------------------------------------
            for (int i = 0; i < columns.Count; i++)
            {
                List<PltContour> column =
                    columns[i];

                // Сортируем внутри колонки по Y
                column.Sort(
                    delegate (
                        PltContour a,
                        PltContour b)
                    {
                        return a.CenterY
                            .CompareTo(b.CenterY);
                    });

                // -------------------------------------------------
                // 0, 2, 4... колонка
                // -------------------------------------------------
                if (i % 2 == 0)
                {
                    // снизу вверх
                    for (int j = 0;
                         j < column.Count;
                         j++)
                    {
                        result.Add(column[j]);
                    }
                }
                else
                {
                    // -------------------------------------------------
                    // 1, 3, 5... колонка
                    // -------------------------------------------------
                    // сверху вниз
                    for (int j = column.Count - 1;
                         j >= 0;
                         j--)
                    {
                        result.Add(column[j]);
                    }
                }
            }

            return result;
        }

        // ---------------------------------------------------------
        // Оценка расстояния между колонками
        // ---------------------------------------------------------
        private double EstimateXStep(
            List<PltContour> contours)
        {
            if (contours.Count < 2)
                return 100;

            List<double> xs =
                new List<double>();

            foreach (PltContour contour in contours)
            {
                xs.Add(contour.CenterX);
            }

            xs.Sort();

            double minStep =
                double.MaxValue;

            for (int i = 1; i < xs.Count; i++)
            {
                double step =
                    xs[i] - xs[i - 1];

                if (step <= 0)
                    continue;

                if (step < minStep)
                    minStep = step;
            }

            if (minStep == double.MaxValue)
                return 100;

            return minStep;
        }

        
    }
}