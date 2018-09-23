using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
namespace Euklidas
{
    class Euklidas1
    {
        Excelmodel excelmodel = new Excelmodel();
        UserInputModel UserInputModel = new UserInputModel();
        ICollection<Excelmodel> excCollection = new List<Excelmodel>(); //  Nepanaudotas
      
        public double EuklidoAtstumas(double[] p1, double[] p2)
        {
           
            double sum = 0.0;
            for(int i = 0; i < p1.Length; i++)
            {
                double u = p1[i] - p2[i];
                sum += u * u;
            }
            return Math.Sqrt(sum);
        }
    }
}
