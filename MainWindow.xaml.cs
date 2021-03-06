﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;


namespace MlabToExcelExport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private SetViewModel GenerateSet()
        {

            double startMIC = 0.125;
            var MIClist = new List<double>();
            for (int i = 1; i <= 10; i++)
            {
                MIClist.Add(startMIC);
                startMIC = startMIC * 2;
            }

            startMIC = 0.03125;
            var ControlMIClist = new List<double>();
            for (int i = 1; i <= 10; i++)
            {
                ControlMIClist.Add(startMIC);
                startMIC = startMIC * 2;
            }

            var MOlist = new List<SetRow>();
            for (int i = 2; i <= 93; i++)
            {
                MOlist.Add(new SetRow
                {
                    Cell = "A" + i,
                    MO = "Oranism " + i,
                    MuseumNumber = "MARAFON-"+(120 + i).ToString(),
                    Number = i
                });
            }

            var ControlMOlist = new List<SetRow>();
            for (int i = 1; i <= 3; i++)
            {
                ControlMOlist.Add(new SetRow
                {
                    Cell = "",
                    MO = "Control Oranism " + i,
                    MuseumNumber = "control" + i,
                    Number = i
                });
            }

            var collection = new ObservableCollection<SetItem>();
            for (int i = 0; i <= 10; i++)
            {
                collection.Add(new SetItem
                {
                    AB = "Antibiotic " + i,
                    Set = "Set Number " + i,
                    Project = "Project " + i,
                    TestMethod = "Метод разведения в агаре",
                    MICList = MIClist,
                    ControlMICList = ControlMIClist,
                    MOList = MOlist,
                    ControlMOList = ControlMOlist
                });
            }

            var obj = new SetViewModel();
            obj.Set = collection;
            return obj;

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            SetViewModel data = GenerateSet();

            MessageBox.Show(ExportToExcel.GetExcelDocumentSetPDF(data, null, 1));
        }
        private void Button_Click_Other(object sender, RoutedEventArgs e)
        {
            SetViewModel data = GenerateSet();
            MessageBox.Show(ExportToExcel.GetExcelDocumentSet(data, null, 2));
        }
        private void Button_Click_Other2(object sender, RoutedEventArgs e)
        {
            if (ExportToExcel.OpenExcelDocument(@"C:\Users\Inspiron\Desktop\maestro_output.xls") == 1)
          {
              MessageBox.Show("File OPened Succesfully");
          }
          else
          {
              MessageBox.Show("Can not open file, probabli incorrect filename");
          }
        }

    }
}
