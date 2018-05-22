using System;
using System.Windows;
using SAPbouiCOM.Framework;
using System.Runtime.InteropServices;
using SAPbobsCOM;
using Microsoft.Win32;
using System.IO;
using System.Security;

namespace ExcelParserToDatabase
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

        }



        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            var program = new ModProgram();
            program.Mains(PathBox.Text, UdoName.Text);
            StatusBar.Text = "Finish successful";
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            // Displays an OpenFileDialog and shows the read/only files.

            OpenFileDialog dlgOpenFile = new OpenFileDialog();
            dlgOpenFile.ShowReadOnly = true;


            dlgOpenFile.ShowDialog();
            {

                // If ReadOnlyChecked is true, uses the OpenFile method to
                // open the file with read/only access.
                string path = null;

                try
                {
                    if (dlgOpenFile.ReadOnlyChecked == true)
                    {
                      
                    }

                    // Otherwise, opens the file with read/write access.
                    else
                    {
                        PathBox.Text = dlgOpenFile.FileName;
                    }
                }
                catch (SecurityException ex)
                {
                    // The user lacks appropriate permissions to read files, discover paths, etc.
                    MessageBox.Show("Security error. Please contact your administrator for details.\n\n" +
                        "Error message: " + ex.Message + "\n\n" +
                        "Details (send to Support):\n\n" + ex.StackTrace
                    );
                }
                catch (Exception ex)
                {
                    // Could not load the image - probably related to Windows file system permissions.
                    MessageBox.Show("Cannot Open File  " + path.Substring(path.LastIndexOf('\\'))
                        + ". You may not have permission to read the file, or " +
                        "it may be corrupt.\n\nReported error: " + ex.Message);
                }
            }
            
        }


    }
}
