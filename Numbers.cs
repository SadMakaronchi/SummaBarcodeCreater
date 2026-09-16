using System;
using System.Linq;
using System.Windows;

namespace SummaBarcodeCreater
{
    public class Numbers
    {
        public long up_number;
        public long down_number;
        public string up_number_barcode;
        public string down_number_barcode;

        
        public Numbers Calc()
        {
            Random rnd = new Random();// генерируем рандомное число
            long barnmbr = rnd.Next(1_000_000, 10_000_000) * 1000L + rnd.Next(0, 1000);// число баркода без индекса стороны резки и контрольного числа

            up_number = long.Parse("6" + barnmbr); // имя plt файла и подпись
            down_number = long.Parse("9" + barnmbr);


            int CalcCheckDigit(long number)// считаем проверочное число
                {
                    int sum = number.ToString()
                                   .Where(char.IsDigit)
                                   .Sum(c => c - '0');
                    return (10 - (sum % 10)) % 10;
                }
            int nm1 = CalcCheckDigit(up_number);
            int nm2 = CalcCheckDigit(down_number);
            up_number_barcode = up_number.ToString() + nm1; //число для генерации баркода
            down_number_barcode = down_number.ToString() + nm2;
            return this;
        }
    }
}
