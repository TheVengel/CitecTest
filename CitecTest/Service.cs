using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CitecTest
{
    internal class Service
    {
        private Repository repository = new Repository();

        public void ReadFile(string filename, int RRK = 0, int Appeal = 0)
        {
            var reader = new StreamReader(filename);
            int valueRRK;
            int valueAppeal;
            if (RRK == 1 && Appeal == 0)
            {
                valueRRK = 1;
                valueAppeal = 0;
            }
            else if (RRK == 0 && Appeal == 1)
            {
                valueRRK = 0;
                valueAppeal = 1;
            }
            else 
            {
                valueRRK = 0;
                valueAppeal = 0;
            }

            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                if (line.Contains("Климов Сергей Александрович"))
                {
                    var indexOfDot1 = line.IndexOf(".");
                    var indexOfDot2 = line.IndexOf(".", indexOfDot1 + 1);
                    
                    var name = line.Substring(28, indexOfDot2 - 27);
                    var administrant = new Administrant(name, valueRRK, valueAppeal);

                    Action(administrant);
                }
                else
                {
                    var items = line.Split(new char[] { ' ' }, 3);
                    var name = items[0] + " " + items[1][0] + "." + items[2][0] + ".";
                    var administrant = new Administrant(name, valueRRK, valueAppeal);

                    Action(administrant);
                }
            }
        }

        public string PrintList()
        {
            var listString = "";
            var list = repository.GetList();
            foreach (var item in list)
            {
                var info = item.GetInfo();
                foreach (var infoItem in info)
                {
                    listString += infoItem;
                    listString += " ";
                }
                listString += Environment.NewLine;

            }
            return listString;
        }

        private void Action(Administrant administrant)
        {
            if (!repository.IsExist(administrant))
                repository.Create(administrant);
            else
                repository.Update(administrant);
        }

        public List<Administrant> GetList()
        {
            return repository.GetList();
        }

        public int GetAllRRK()
        {
            int RRK = 0;
            foreach (var administrant in repository.GetList())
                RRK += administrant.DocsRRK;
            return RRK;
        }

        public int GetAllAppeal()
        {
            int Appeal = 0;
            foreach (var administrant in repository.GetList())
                Appeal += administrant.DocsAppeal;
            return Appeal;
        }

        public int GetSumDocs()
        {
            return GetAllRRK() + GetAllAppeal();
        }
    }
}
