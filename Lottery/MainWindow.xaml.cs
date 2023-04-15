using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace Lottery
{
    public partial class MainWindow : Window
    {
        PersonModel model;

        List<Person> persons = new();
        List<Person> personsRemove = new();

        string path = @"D:\table.xlsx";
        string pathRemove = @"D:\remove.xlsx";
        public MainWindow()
        {
            InitializeComponent();

            this.WindowState = WindowState.Maximized;

            persons = Read(path, 2);

            if (File.Exists(pathRemove))
            {
                personsRemove = Read(pathRemove, 1);
                personsRemove = personsRemove?.Distinct(new PersonComparer()).ToList();
                persons = persons?.Except(personsRemove!, new PersonComparer()).ToList();
            }

            model = new PersonModel(persons);
            DataContext = model;

        }

        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
            Random random = new();
            var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(50) };
            timer.Tick += timerTick;
            timer.Start();

            var timer10 = new DispatcherTimer { Interval = TimeSpan.FromSeconds(10) };
            timer10.Tick += (sender, e) =>
            {
                timer.Stop();
                timer10.Stop();

                //MessageBox.Show(model.CurrentPerson.LastName);
                personsRemove.Add(model.CurrentPerson);
                model.Persons.Remove(model.CurrentPerson);



                return;
            };
            timer10.Start();

            void timerTick(object sender, EventArgs e)
            {
                int index = random.Next(0, model.Persons.Count);
                model.CurrentPerson = model.Persons[index];
            }
        }

        List<Person> Read(string path, int start)
        {
            List<Person> ps = new();

            FileInfo fileInfo = new FileInfo(path);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage package = new ExcelPackage(fileInfo);
            ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

            int rows;
            if (worksheet.Dimension.Rows != null)
                rows = worksheet.Dimension.Rows;
            else
                return null;

            int columns = worksheet.Dimension.Columns;

            for (int i = start; i <= rows; i++)
            {
                Person person = new();
                for (int j = 1; j <= columns; j++)
                {
                    string content = "";

                    if (worksheet.Cells[i, j].Value != null)
                        content = worksheet.Cells[i, j].Value.ToString();

                    switch (j)
                    {
                        case 1: person.Time = content; break;
                        case 2: person.FirstName = content; break;
                        case 3: person.LastName = content; break;
                        case 4: person.Phone = content; break;
                        case 5: person.Email = content; break;
                    }
                }
                ps.Add(person);
            }
            return ps;
        }

        void Write()
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets.Add("PemovePersons");

                int row = 1;
                foreach (var person in personsRemove)
                {
                    sheet.Cells[row, 1].Value = person.Time;
                    sheet.Cells[row, 2].Value = person.FirstName;
                    sheet.Cells[row, 3].Value = person.LastName;
                    sheet.Cells[row, 4].Value = person.Phone;
                    sheet.Cells[row, 5].Value = person.Email;
                    row++;
                }

                sheet.Cells.AutoFitColumns();
                package.SaveAs(new FileInfo(pathRemove));
            }
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            Write();
        }
    }
}
