using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Threading;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

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
        //private static List<Course> listCourses = new List<Course>();
        private ObservableCollection<Course> courseCollection = new ObservableCollection<Course>();
        private BackgroundWorker worker = new BackgroundWorker();

        public MainWindow()
        {
            InitializeComponent();

            startCol = Convert.ToInt32(startColTxt.Text);
            totalCols = Convert.ToInt32(totalColsTxt.Text);
            startDatePicker.SelectedDate = DateTime.Today;
            endDatePicker.SelectedDate = DateTime.Today;
            lblPath.Text = "Ningún archivo seleccionado";
            errorImage.Visibility = Visibility.Visible;
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Archivo Excel (*xlsx)|*.xlsx";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            if (openFileDialog.ShowDialog() == true)
                filePath = openFileDialog.FileName;
                    if(filePath != null)
            {
                btnDoAsynchronousCalculation.IsEnabled = true;
                lblPath.Text = filePath;
                ControlImages(true);
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
            lbResults.Items.Refresh();
            lblPath.Text = "Leyendo archivo...";
            ControlBtns(false);
            okImage.Visibility = Visibility.Collapsed;
            errorImage.Visibility = Visibility.Collapsed;

            worker.WorkerSupportsCancellation = true;
            worker.WorkerReportsProgress = true;
            worker.DoWork += worker_DoWork;
            worker.ProgressChanged += worker_ProgressChanged;
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;
            worker.RunWorkerAsync(Convert.ToInt32(totalColsTxt.Text));
        }

        private void btnCancelCalculation_Click(object sender, RoutedEventArgs e)
        {
            worker.CancelAsync();
            lblPath.Text = "Cancelando...";
            Thread.Sleep(2000);
            ControlBtns(true);
        }

        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            int max = (int)e.Argument;
            int result = 0; // Contador            
            string matrixValue;
            int startRow = 7;            

            for (int col = startCol; col <= max; col++)
            {
                if (worker.CancellationPending == true)
                {
                    e.Cancel = true;
                    return;
                }

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
                courseCollection.Add(new Course() { Name = mV[0], Part = mV[1], DateCount = mV[2], Total = mV[3]});
                lbResults.ItemsSource = courseCollection;
            }
        }

        void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if(e.Cancelled)
            {
                lblPath.Text = "Cancelado por el usuario";
                ControlImages(false);
            }
            else
            {
                lblPath.Text = "Completado! " + e.Result + " Registros leidos";
                btnCreateReport.IsEnabled = true;
                ControlImages(true);
            }

            ControlBtns(true);
        }

        private void btnCreateReport_Click(object sender, RoutedEventArgs e)
        {
            FileExcel.CreateReport(courseCollection);
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

        private void ControlBtns(bool c)
        {
            if(c)
            {
                btnCancelCalculation.Visibility = Visibility.Collapsed;
                btnDoAsynchronousCalculation.Visibility = Visibility.Visible;
            }
            else
            {
                btnCancelCalculation.Visibility = Visibility.Visible;
                btnDoAsynchronousCalculation.Visibility = Visibility.Collapsed;
            }            
        }

        private void ControlImages(bool c)
        {
            if (c)
            {
                okImage.Visibility = Visibility.Visible;
                errorImage.Visibility = Visibility.Collapsed;
            }
            else
            {
                okImage.Visibility = Visibility.Collapsed;
                errorImage.Visibility = Visibility.Visible;
            }
        }
    }    
}
