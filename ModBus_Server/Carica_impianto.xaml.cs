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
using System.IO;

namespace ModBus_Server
{
    /// <summary>
    /// Interaction logic for Carica_impianto.xaml
    /// </summary>
    public partial class Carica_impianto : Window
    {
        public string path { get; set; }

        public Carica_impianto(String defaultPath)
        {
            InitializeComponent();

            String[] subFolders = Directory.GetDirectories("Json\\");

            comboBoxImpianto.Items.Clear();

            for (int i = 0; i < subFolders.Length; i++)
            {
                comboBoxImpianto.Items.Add(subFolders[i].Substring(subFolders[i].IndexOf('\\') + 1));
            }

            comboBoxImpianto.SelectedItem = defaultPath;
        }

        private void buttonOpen_Click(object sender, RoutedEventArgs e)
        {
            //path = comboBoxImpianto.SelectedValue.ToString().Substring(comboBoxImpianto.SelectedValue.ToString().IndexOf(' '), comboBoxImpianto.SelectedValue.ToString().Length - comboBoxImpianto.SelectedValue.ToString().IndexOf(' '));
            path = comboBoxImpianto.SelectedValue.ToString();

            // debug
            Console.WriteLine("path: " + path);

            this.DialogResult = true;
        }

        private void buttonAnnulla_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
        }
    }
}
