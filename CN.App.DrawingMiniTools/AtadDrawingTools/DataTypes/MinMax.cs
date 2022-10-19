using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ATADDrawingTools.DataTypes
{
    class MinMax
    {
        public int FindMax(int num1, int num2)
        {
            /* khai bao bien cuc bo */
            int result;
            if (num1 > num2)
                result = num1;
            else
                result = num2;
            return result;
        }
        public int FindMin(int num1, int num2)
        {
            /* khai bao bien cuc bo */
            int result;
            if (num1 < num2)
                result = num1;
            else
                result = num2;
            return result;
        }
    }
}
