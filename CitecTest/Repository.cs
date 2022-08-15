using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CitecTest
{
    internal class Repository
    {
        private List<Administrant> Administrants { get; set; }

        public Repository()
        {
            this.Administrants = new List<Administrant>();
        }

        public void Create(Administrant administrant)
        {
            Administrants.Add(administrant);
        }

        public void Update(Administrant administrant)
        {
            foreach (var item in Administrants)
            {
                if (item.Equals(administrant))
                {
                    item.DocsRRK += administrant.DocsRRK;
                    item.DocsAppeal += administrant.DocsAppeal;
                    break;
                }
            }
        }

        public bool IsExist(Administrant administrant)
        {
            foreach (var item in Administrants)
            {
                if (item.Equals(administrant))
                    return true;
            }
            return false;
        }

        public int? Read(Administrant administrant)
        {
            foreach (var item in Administrants)
            {
                if (item.Equals(administrant))
                    return this.Administrants.IndexOf(item);
            }
            return null;
        }

        public List<Administrant> GetList()
        {
            return this.Administrants;
        }
    }
}
