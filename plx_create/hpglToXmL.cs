using System;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;

public class HpglToXml
{
    public static XElement Convert(string hpgl)
    {
        XNamespace ns = "urn:schemas-colorgate-com:cutjobinformation";

        XElement curves = new XElement(ns + "Curves");

        MatchCollection matches = Regex.Matches(hpgl,@"(?i)(PD|PU)\s*(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)"
        );

        foreach (Match match in matches)
        {
            string command = match.Groups[1].Value.ToUpperInvariant();

            double x = double.Parse(
                match.Groups[2].Value,
                CultureInfo.InvariantCulture
            );

            double y = double.Parse(
                match.Groups[3].Value,
                CultureInfo.InvariantCulture
            );

            double xMm = x * 0.025;
            double yMm = y * 0.025;

            XElement curve = new XElement(
                ns + "Curve",
                new XAttribute(
                    "Type",
                    command == "PD" ? "Line" : "Move"
                ),
                new XElement(
                    ns + "Coordinate",
                    new XAttribute(
                        "X",
                        xMm.ToString("0.000", CultureInfo.InvariantCulture)
                    ),
                    new XAttribute(
                        "Y",
                        yMm.ToString("0.000", CultureInfo.InvariantCulture)
                    )
                )
            );

            curves.Add(curve);
        }

        return curves;
    }
}