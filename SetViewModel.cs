using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MlabToExcelExport
{
    public class SetViewModel
    {
        public ObservableCollection<SetItem> Set { get; set; }
    }

    public class SetItem
    {
        public string AB { get; set; }
        public string Project { get; set; }
        public string TestMethod { get; set; }
        public string Set { get; set; }
        public List<double> MICList { get; set; }

        public List<double> ControlMICList { get; set; }

        public List<SetRow> MOList { get; set; }
        public List<SetRow> ControlMOList { get; set; }
    }

    public class SetRow
    {
        public string Cell { get; set; }
        public int Number { get; set; }
        public string MuseumNumber { get; set; }
        public string MO { get; set; }


    }
}
