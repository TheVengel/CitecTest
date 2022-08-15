using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CitecTest
{
    internal class Administrant
    {
        public string Name { get; set; }
        public int DocsRRK { get; set; }

        public int DocsAppeal { get; set; }

        public Administrant(string name, int docsRRK = 0, int docsAppeal = 0)
        {
            this.Name = name;
            this.DocsRRK = docsRRK;
            this.DocsAppeal = docsAppeal;
        }

        public override bool Equals(object? obj)
        {
            if (obj is Administrant administrant)
                return this.Name == administrant.Name;
            return false;
        }

        public List<string> GetInfo()
        {
            return new List<string>
            {
                this.Name,
                this.DocsRRK.ToString(),
                this.DocsAppeal.ToString()
            };
        }
    }
}
