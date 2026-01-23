using Corel.Interop.VGCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Shapes;
using corel = Corel.Interop.VGCore;
using Shape = Corel.Interop.VGCore.Shape;


namespace SummaMetki
{
    public class BarcodeCreater
    {
        public corel.Application corelApp;
        public void Init(corel.Application app)
        {
            corelApp = app;
        }
        public void Create(Layer brk, Shape met3,Shape met4, string bar1, string bar2)
        {

            string bn;
            string bin;
            if (met4 == null)
            {
                bin = DecimalToPostnet(bar1);
                
            }
            else
            {
                bin = DecimalToPostnet(bar2);
            }

                string DecimalToPostnet(string dec)
                {
                    bin = "";
                    bn = "";
                    for (int n = 0; n < dec.Length; n++)
                    {
                        bn = Conv(dec[n]);
                        bin = bin + bn;

                    }
                    return bin;
                }
                string Conv(char n)
                {
                    if (n == '0') 
                    { 
                        bn = "11000";
                    }
                    if (n == '1') 
                    { 
                        bn = "00011"; 
                    } 
                    if (n == '2') 
                    { 
                        bn = "00101"; 
                    } 
                    if (n == '3') 
                    { 
                        bn = "00110"; 
                    }
                    if (n == '4') 
                    { 
                        bn = "01001"; 
                    }
                    if (n == '5')
                    { 
                        bn = "01010"; 
                    } 
                    if (n == '6') 
                    { 
                        bn = "01100"; 
                    } 
                    if (n == '7') 
                    { 
                        bn = "10001"; 
                    } 
                    if (n == '8') 
                    { 
                        bn = "10010"; 
                    } 
                    if (n == '9') 
                    { 
                        bn = "10100"; 
                    } 
                    return bn; 
                }
                    
            
            

                

                
                bin = "1" + bin + "1";
            
            for (int i = 0; i < bin.Length; i++)
                {
                    char bit = bin[i];
                    if (bit == '1')
                    {
                        PaintOne(i);
                    }
                    else
                    {
                        PaintZero(i);
                    }
                }
                void PaintOne(int i)
                {
                    var one = brk.CreateRectangle(0, 9.53, 1.52, 0);
                    one.ConvertToCurves();
                    one.Style.StringAssign(@"{""outline"":{""width"":""0""}}");
                    one.Fill.UniformColor.CMYKAssign(0, 0, 0, 100);
                Move(one, i);
                }
                void PaintZero(int i)
                {
                    var zero = brk.CreateRectangle(0, 3.81, 1.52, 0);
                    zero.ConvertToCurves();
                    zero.Style.StringAssign(@"{""outline"":{""width"":""0""}}");
                    zero.Fill.UniformColor.CMYKAssign(0, 0, 0, 100);
                Move(zero,i);
                }
            void Move(Shape bar,int i)
            {
                if (met4==null)
                {
                    bar.LeftX = met3.RightX - 215.63 + 3.51 * i;
                    bar.BottomY = met3.TopY;
                }
                else
                {
                    bar.LeftX = met4.LeftX + 215.63 - 3.51 * i;
                    bar.BottomY = met4.TopY;
                }
            }


        }
    }
}
