using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Input;

namespace BgWorkerApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string filePath;
        private int startCol;
        private int totalCols;
        private DateTime startDate;
        private DateTime endDate;
        private List<Course> listCourses = new List<Course>();


        public MainWindow()
        {
            InitializeComponent();

            startCol = Convert.ToInt32(startColTxt.Text);
            totalCols = Convert.ToInt32(totalColsTxt.Text);
            startDatePicker.SelectedDate = DateTime.Today;
            endDatePicker.SelectedDate = DateTime.Today;
            lblPath.Text = "Ningún archivo seleccionado";
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
                filePath = openFileDialog.FileName;
                    if(filePath != null)
            {
                btnDoAsynchronousCalculation.IsEnabled = true;
                lblPath.Text = filePath;
            }                        
        }

        private void btnDoAsynchronousCalculation_Click(object sender, RoutedEventArgs e)
        {            
            startDate = startDatePicker.SelectedDate.Value;
            endDate = endDatePicker.SelectedDate.Value;
            startCol = Convert.ToInt32(startColTxt.Text);
            totalCols = Convert.ToInt32(totalColsTxt.Text);
            pbCalculationProgress.Value = 0;
            lbResults.Items.Clear();
            lblPath.Text = "Leyendo archivo...";

            BackgroundWorker worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.DoWork += worker_DoWork;
            worker.ProgressChanged += worker_ProgressChanged;
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;
            worker.RunWorkerAsync(Convert.ToInt32(totalColsTxt.Text));
        }

        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            int max = (int)e.Argument;
            int result = 0; // Contador            
            string matrixValue;
            int startRow = 7;

            for (int col = startCol; col <= max; col++)
            {
                int progressPercentage = Convert.ToInt32(((double)col / max) * 100);
                
                matrixValue = FileExcel.Read(filePath, startRow, col, startDate, endDate);

                if (matrixValue != null) { result++; }
                
                (sender as BackgroundWorker).ReportProgress(progressPercentage, matrixValue);
                System.Threading.Thread.Sleep(1);
            }
            e.Result = result;
        }

        void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pbCalculationProgress.Value = e.ProgressPercentage;
            lblPath.Text = "Leyendo archivo...  " + e.ProgressPercentage + "%";
            {
                string[] mV = e.UserState.ToString().Split('^');
                listCourses.Add(new Course() { Name = mV[0], Part = mV[1], DateCount = mV[2], Total = mV[3]});
                lbResults.ItemsSource = listCourses;
            }
                //lbResults.Items.Add(e.UserState);
        }

        void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            lblPath.Text = "Completado! " + e.Result + " Registros leidos";
            btnCreateReport.IsEnabled = true;
        }

        private void btnCreateReport_Click(object sender, RoutedEventArgs e)
        {
            FileExcel.CreateReport(listCourses);
        }

        private void StackPanel_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (Mouse.LeftButton == MouseButtonState.Pressed)
                this.DragMove();
        }

        private void Image_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Window parentWindow = Window.GetWindow(this);            
            App.Current.Shutdown();
        }
    }    
}
