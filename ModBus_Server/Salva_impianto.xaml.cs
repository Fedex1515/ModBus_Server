using System;
using System.Collections.Generic;
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
using System.Windows.Shapes;
using Microsoft.Win32;

namespace ModBus_Server
{
    /// <summary>
    /// Interaction logic for Salva_impianto.xaml
    /// </summary>
    public partial class Salva_impianto : Window
    {
        public string path { get; set; }

        public Salva_impianto()
        {
            InitializeComponent();
        }

        private void buttonOpen_Click(object sender, RoutedEventArgs e)
        {
            path = textBoxImpianto.Text.ToString();
            this.DialogResult = true;
        }

        private void buttonAnnulla_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
        }

    }
}
