using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MlabToExcelExport
{
    public class SetViewModel
    {
        public ObservableCollection<SetItem> Set
        {
            get
            {
                ObservableCollection<SetItem> collection = new ObservableCollection<SetItem>();
                for (int i = 0; i <= 10; i++)
                {
                    collection.Add(new SetItem
                    {
                        AB = "Antibiotic " + i,
                        Set = "Set Number " + i ,
                        Project = "Project " + i,
                        TestMethod = "Метод разведения в агаре"
                    });
                }
                return collection;
            }
        }
    }
}
