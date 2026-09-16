using Corel.Interop.VGCore;
using System;

namespace SummaBarcodeCreater
{
    public class OPOS_marks
    {
        Shape mark_left;
        Shape mark_right;
        Shape mark_top;
        Shape mark_bottom;
        ShapeRange marks;
        public Settings settings;
        public LayerAndSearchContour layerAndSearchContour;
        public Define_per define_per;
        public double step_y;
        public int metk_y;
        public OPOS_marks Create_marks(LayerAndSearchContour layerAndSearchContour, Settings settings,Define_per define_per)
        {
            this.define_per = define_per;
            this.settings = settings;
            this.layerAndSearchContour = layerAndSearchContour;
            bool barc2 = settings.barc2;
            Layer metk_sum = layerAndSearchContour.metk_sum;
            ShapeRange r = define_per.allShapes;
            // считаем шаг и количество меток OPOS
            double b2 = 0;
            if (barc2 == true)
            {
                b2 = 15.53;
            }
            metk_y = (int)Math.Ceiling((r.SizeHeight + 15.53 + b2) / 400);
            step_y = (r.SizeHeight + 15.53 + b2) / metk_y;
            if (step_y > 410)
            {
                metk_y = (int)Math.Ceiling((r.SizeHeight + 15.53 + b2) / 400 + 1);
                step_y = (r.SizeHeight + 15.53 + b2) / metk_y;
            }


            for (int q = 0; q < metk_y + 1; q++) // отрисовка меток OPOS
            {
                mark_left = metk_sum.CreateRectangle(
                                                     r.LeftX - 13, 
                                                     r.BottomY - 12.53 + q * step_y, 
                                                     r.LeftX - 10, 
                                                     r.BottomY - 15.53 + q * step_y);
                mark_right = metk_sum.CreateRectangle(
                                                      r.RightX + 13, 
                                                      r.BottomY - 12.53 + q * step_y, 
                                                      r.RightX + 10, 
                                                      r.BottomY - 15.53 + q * step_y);
            }
            mark_bottom = metk_sum.CreateRectangle(
                                                   metk_sum.Shapes.All().LeftX + 13, 
                                                   metk_sum.Shapes.All().BottomY, 
                                                   metk_sum.Shapes.All().RightX - 13, 
                                                   metk_sum.Shapes.All().BottomY + 3); // отрисовка нижней метки OPOS_XY
            if (barc2 == true)
            {
                mark_top = metk_sum.CreateRectangle(
                                                    metk_sum.Shapes.All().LeftX + 13, 
                                                    metk_sum.Shapes.All().TopY, 
                                                    metk_sum.Shapes.All().RightX - 13, 
                                                    metk_sum.Shapes.All().TopY - 3); // отрисовка верхней метки OPOS_XY
            }
            marks = metk_sum.Shapes.All();
            SetMarkStyle(marks);
            define_per.allShapes.AddRange(marks);
            return this;
        }

        
        public void SetMarkStyle(ShapeRange marks)
        {
            foreach (Shape mark in marks)
            {
                mark.Fill.UniformColor.CMYKAssign(0, 0, 0, 100);
                mark.Style.StringAssign(@"{""outline"":{""width"":""0""}}");
            }
        }
    }
}
