using iText.Kernel.Colors;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using System;
using System.Collections.Generic;

namespace SummaBarcodeCreater
{
    public class PdfContourExtractor
    {
        public class Contour
        {
            public List<(double X, double Y)> Points { get; set; }
                = new List<(double X, double Y)>();

            public bool IsClosed { get; set; }
        }

        private class PathListener : IEventListener
        {
            private readonly string spotColorName;

            public List<Contour> Contours { get; } =
                new List<Contour>();

            public PathListener(string spotColorName)
            {
                this.spotColorName = spotColorName;
            }

            public void EventOccurred(IEventData data,EventType type)
            {
                if (type != EventType.RENDER_PATH)
                {
                    return;
                }

                var info = (PathRenderInfo)data;

                // Нас интересует только обводка.
                if ((info.GetOperation() & PathRenderInfo.STROKE) == 0)
                {
                    return;
                }
                // Отбрасываем пути с другим цветом.
                if (!IsTargetSpotColor(info.GetStrokeColor(), spotColorName))
                {
                    return;
                }

                Path path = info.GetPath();

                foreach (Subpath subpath in path.GetSubpaths())
                {
                    if (subpath.IsEmpty())
                    {  
                        continue; 
                    }
                    // Получаем ломаную аппроксимацию
                    // линий и кривых Безье.
                    var approximation = subpath.GetPiecewiseLinearApproximation();

                    var contour = new Contour
                    {
                        IsClosed = subpath.IsClosed()
                    };

                    foreach (Point point in approximation)
                    {
                        contour.Points.Add((point.GetX(),point.GetY()));
                    }

                    if (contour.Points.Count > 0)
                    {
                        Contours.Add(contour);
                    }
                }
            }

            public ICollection<EventType> GetSupportedEvents()
            {
                return new HashSet<EventType>
            {
                EventType.RENDER_PATH
            };
            }

            private static bool IsTargetSpotColor(Color color,string targetName)
            {
                if (color == null)
                {
                    return false;
                }
                var colorSpace = color.GetColorSpace();
                var colorSpaceObject = colorSpace.GetPdfObject();

                // Separation имеет структуру:
                // [/Separation /ИмяЦвета /AlternateSpace /TintTransform]
                if (!(colorSpaceObject is PdfArray array))
                {
                    return false;
                }
                if (array.Size() < 2)
                {
                    return false;
                }
                PdfObject spaceType = array.Get(0);
                PdfObject colorName = array.Get(1);

                if (!(spaceType is PdfName typeName) || !PdfName.Separation.Equals(typeName))
                {
                    return false;
                }
                if (!(colorName is PdfName spotName))
                {
                    return false;
                }
                return string.Equals(spotName.GetValue(),targetName,StringComparison.Ordinal);
            }
        }

        public static List<Contour> Extract(string pdfPath,string spotColorName,int pageNumber = 1)
        {
            using (var reader = new PdfReader(pdfPath))
            using (var pdf = new PdfDocument(reader))
            {
                if (pageNumber < 1 ||pageNumber > pdf.GetNumberOfPages())
                {
                    throw new ArgumentOutOfRangeException(
                        nameof(pageNumber));
                }

                var listener = new PathListener(spotColorName);
                var processor = new PdfCanvasProcessor(listener);

                processor.ProcessPageContent(pdf.GetPage(pageNumber));

                return listener.Contours;
            }
        }
    }
}
