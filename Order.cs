
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace SummaBarcodeCreater
{
    public class Order
    {
        // Одна точка PLT.
        private struct PltPoint
        {
            public double X;
            public double Y;

            public PltPoint(double x, double y)
            {
                X = x;
                Y = y;
            }
        }

        // Один контур и все его исходные команды.
        private class PltContour
        {
            public List<string> Commands { get; } = new List<string>();

            public PltPoint Start;
            public PltPoint End;

            public int OriginalIndex;

            public PltContour(int index)
            {
                OriginalIndex = index;
            }

            public double DistanceSquared(PltPoint point)
            {
                double dx = Start.X - point.X;
                double dy = Start.Y - point.Y;

                return dx * dx + dy * dy;
            }
        }

        // Поддерживает команды формата:
        // PU4194,10499;
        // PD4200,10500;
        // PU 4194,10499;
        private static readonly Regex CoordinateCommand = new Regex(
            @"^\s*(PU|PD)\s*(-?\d+)\s*,\s*(-?\d+)\s*;\s*$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// Основной метод.
        /// Принимает PLT-текст и возвращает контуры,
        /// отсортированные по ближайшему следующему контуру.
        /// </summary>
        public string Process(string plt)
        {
            return Process(plt, 0, 0);
        }

        /// <summary>
        /// Позволяет задать начальную позицию ножа.
        /// По умолчанию используется координата (0, 0).
        /// </summary>
        public string Process(string plt, double initialX, double initialY)
        {
            if (string.IsNullOrWhiteSpace(plt))
                return plt;

            List<string> prefix;
            List<PltContour> contours;

            Parse(plt, out prefix, out contours);

            if (contours.Count <= 1)
                return plt;

            List<PltContour> ordered = SortNearest(
                contours,
                new PltPoint(initialX, initialY));

            return BuildOutput(prefix, ordered);
        }

        /// <summary>
        /// Разбирает PLT на команды до первого контура
        /// и отдельные контуры, начинающиеся с PUx,y;.
        /// </summary>
        private void Parse(
            string plt,
            out List<string> prefix,
            out List<PltContour> contours)
        {
            prefix = new List<string>();
            contours = new List<PltContour>();

            string[] lines = plt.Split(
                new[] { "\r\n", "\n", "\r" },
                StringSplitOptions.None);

            PltContour current = null;

            foreach (string line in lines)
            {
                Match match = CoordinateCommand.Match(line);

                bool isCoordinateCommand = match.Success;
                bool isPenUpWithCoordinates =
                    isCoordinateCommand &&
                    match.Groups[1].Value.Equals(
                        "PU",
                        StringComparison.OrdinalIgnoreCase);

                if (isPenUpWithCoordinates)
                {
                    // Каждая PU с координатами начинает новый контур.
                    current = new PltContour(contours.Count);

                    current.Start = new PltPoint(
                        ParseCoordinate(match.Groups[2].Value),
                        ParseCoordinate(match.Groups[3].Value));

                    current.End = current.Start;
                    current.Commands.Add(line);

                    contours.Add(current);
                    continue;
                }

                if (current == null)
                {
                    // Команды до первого контура сохраняем в начале файла.
                    prefix.Add(line);
                    continue;
                }

                current.Commands.Add(line);

                if (isCoordinateCommand &&
                    match.Groups[1].Value.Equals(
                        "PD",
                        StringComparison.OrdinalIgnoreCase))
                {
                    current.End = new PltPoint(
                        ParseCoordinate(match.Groups[2].Value),
                        ParseCoordinate(match.Groups[3].Value));
                }
            }
        }

        private double ParseCoordinate(string value)
        {
            return double.Parse(value, CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Жадная сортировка ближайшего соседа.
        /// На каждом шаге выбирается контур,
        /// стартовая точка которого ближе всего к текущей позиции.
        /// </summary>
        private List<PltContour> SortNearest(
            List<PltContour> contours,
            PltPoint initialPosition)
        {
            var remaining = new List<PltContour>(contours);
            var result = new List<PltContour>(contours.Count);

            PltPoint currentPosition = initialPosition;

            while (remaining.Count > 0)
            {
                int nearestIndex = 0;
                double nearestDistance = double.MaxValue;

                for (int i = 0; i < remaining.Count; i++)
                {
                    double distance =
                        remaining[i].DistanceSquared(currentPosition);

                    if (distance < nearestDistance)
                    {
                        nearestDistance = distance;
                        nearestIndex = i;
                    }
                }

                PltContour nearest = remaining[nearestIndex];

                result.Add(nearest);
                remaining.RemoveAt(nearestIndex);

                // После резки текущая позиция — конец выбранного контура.
                currentPosition = nearest.End;
            }

            return result;
        }

        /// <summary>
        /// Собирает PLT обратно, сохраняя команды контуров
        /// и исходный префикс файла.
        /// </summary>
        private string BuildOutput(
            List<string> prefix,
            List<PltContour> ordered)
        {
            var output = new StringBuilder();

            foreach (string line in prefix)
                output.AppendLine(line);

            foreach (PltContour contour in ordered)
            {
                foreach (string line in contour.Commands)
                    output.AppendLine(line);
            }

            return output.ToString();
        }
    }
}
