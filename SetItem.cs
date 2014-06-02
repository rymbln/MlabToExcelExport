using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MlabToExcelExport
{
    public class SetItem
    {
        public string AB { get; set; }
        public string Project { get; set; }
        public string TestMethod { get; set; }
        public string Set { get; set; }
        public List<double> MICList
        {
            get
            {
                double startMIC = 0.125;
                List<double> list = new List<double>();
                for (int i = 1; i <= 10; i++)
                {
                    list.Add(startMIC);
                    startMIC = startMIC*2;
                }
                return list;
            }
        }

        public List<SetRow> MOList
        {
            get
            {
                List<SetRow> list = new List<SetRow>();
                for (int i = 1; i <= 40; i++)
                {
                    list.Add(new SetRow
                    {
                        Cell = "A" + i,
                        MO = "Oranism " + i,
                        MuseumNumber = (120 + i).ToString(),
                        Number = i
                    });
                }
                return list;
            }
        }
        public List<SetRow> ControlMOList
        {
            get
            {
                List<SetRow> list = new List<SetRow>();
                for (int i = 1; i <= 3; i++)
                {
                    list.Add(new SetRow
                    {
                        Cell = "",
                        MO = "Control Oranism " + i,
                       MuseumNumber = "control" + i,
                        Number = i
                    });
                }
                return list;
            }
        }
    }
}
