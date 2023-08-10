using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutodeskTestLab
{
    internal class BlockEnt
    {
        string id;
        double weight;
        int diameter;
        string fullNameTemplate;

        public BlockEnt(string id, double weight, int diameter, string fullNameTemplate)
        {
            this.id = id;
            this.weight = weight;
            this.diameter = diameter;
            this.fullNameTemplate = fullNameTemplate;
        }

        public string Id { get => id; set => id = value; }
        public double Weight { get => weight; set => weight = value; }
        public int Diameter { get => diameter; set => diameter = value; }
        public string FullNameTemplate { get => fullNameTemplate; set => fullNameTemplate = value; }
    }
}
