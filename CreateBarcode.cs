using Corel.Interop.VGCore;

namespace SummaBarcodeCreater
{
    public class CreateBarcode
    {
        
        public LayerAndSearchContour layerAndSearchContour;
        public Numbers numbers;
        public Settings settings;
        public ShapeRange barcodes;
        public Define_per define_per;
        public Shape barcode1;
        public void Create(LayerAndSearchContour layerAndSearchContour, Numbers numbers, Settings settings,Define_per define_Per)
        {
            this.layerAndSearchContour = layerAndSearchContour;
            this.numbers = numbers;
            this.settings = settings;
            this.define_per = define_Per;
            Layer brk = layerAndSearchContour.brk;
            Layer metk_sum = layerAndSearchContour.metk_sum;
            bool barc2 = settings.barc2;
            string number_barcode = numbers.down_number_barcode;

            ShapeRange barshape = Creating(number_barcode, brk);
            barshape.RightX = metk_sum.Shapes.All().RightX - 13;
            barshape.BottomY = metk_sum.Shapes.All().BottomY + 3;
            Shape barcode1 = barshape.Combine();
            Shape podp1 = brk.CreateArtisticText(barcode1.RightX - 245, 
                                                barcode1.BottomY + 3, 
                                                numbers.down_number.ToString(), 
                                                cdrTextLanguage.cdrEnglishUS, 
                                                cdrTextCharSet.cdrCharSetDefault,
                                                "Arial", 12, cdrTriState.cdrFalse, 
                                                cdrTriState.cdrFalse);
            podp1.ConvertToCurves();
            podp1.Style.StringAssign(@"{""outline"":{""width"":""0""}}");
            podp1.Fill.UniformColor.CMYKAssign(0, 0, 0, 40);
            define_per.allShapes.Add(podp1);
            define_per.allShapes.Add(barcode1);
            if (barc2 == true)
            {
                number_barcode = numbers.up_number_barcode;
                barshape = Creating(number_barcode, brk);
                barshape.Rotate(180);
                barshape.LeftX = metk_sum.Shapes.All().LeftX + 13;
                barshape.TopY = metk_sum.Shapes.All().TopY - 3;
                Shape barcode2 = barshape.Combine();
                Shape podp2 = brk.CreateArtisticText(barcode2.LeftX + 220,
                                               barcode2.TopY - 6,
                                               numbers.up_number.ToString(),
                                               cdrTextLanguage.cdrEnglishUS,
                                               cdrTextCharSet.cdrCharSetDefault,
                                               "Arial", 12, cdrTriState.cdrFalse,
                                               cdrTriState.cdrFalse);
                podp2.ConvertToCurves();
                podp2.Style.StringAssign(@"{""outline"":{""width"":""0""}}");
                podp2.Fill.UniformColor.CMYKAssign(0, 0, 0, 40);
                define_per.allShapes.Add(podp2);
                define_per.allShapes.Add(barcode2);
            }
        }



        public ShapeRange Creating(string number_barcode, Layer brk)
        {
            string bin_num;
            ShapeRange barshape = define_per.corelDoc.CreateShapeRangeFromArray(); 
            var convertToBin = new ConvertToBin();
            bin_num = convertToBin.DecimalToPostnet(number_barcode);
            for (int i = 0; i < bin_num.Length; i++)
            {
                char bit = bin_num[i];
                Shape bar;
                if (bit == '1')
                {
                    bar = brk.CreateRectangle(0, 9.53, 1.52, 0);
                }
                else
                {
                    bar = brk.CreateRectangle(0, 3.81, 1.52, 0);
                }
                bar.ConvertToCurves();
                bar.Style.StringAssign(@"{""outline"":{""width"":""0""}}");
                bar.Fill.UniformColor.CMYKAssign(0, 0, 0, 100);
                bar.LeftX = 0 + 3.51 * i;
                barshape.Add(bar);
            }  
             return barshape;
        }

    }
        public class ConvertToBin
        {
            string bin;
            string bn;
            public string DecimalToPostnet(string number_barcode)
            {
                bin = "";
                bn = "";
                for (int n = 0; n < number_barcode.Length; n++)
                {
                    char i = number_barcode[n];
                    bn = Conv(i);
                    bin = bin + bn;
                }
                bin = "1" + bin + "1"; //добавляем начальные и конечные символы штрихкода
                return bin;
            }
            public string Conv(char i)
            {
                switch (i)
                {
                    case '0': return "11000";
                    case '1': return "00011";
                    case '2': return "00101";
                    case '3': return "00110";
                    case '4': return "01001";
                    case '5': return "01010";
                    case '6': return "01100";
                    case '7': return "10001";
                    case '8': return "10010";
                    case '9': return "10100";
                }
                return bn;
            }

        }
}
