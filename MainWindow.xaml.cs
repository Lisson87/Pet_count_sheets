using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection.Metadata;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
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
using System.Windows.Threading;
using Microsoft.Win32;
using PdfSharpCore;
using PdfSharpCore.Pdf.IO;
using static Pet_count_sheets.MainWindow;


namespace Pet_count_sheets
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        DataTable table = new DataTable();
        List<string> FilePaths = new List<string>();    // содержит список файлов pdf (после шага 1)
        List<List<Page_size>> L_LPages = new List<List<Page_size>>();         // содержит список со списками страниц с форматами (после шага 2)
        List<string> colNames = new List<string>();     // СПИСОК КОЛОНОК в зависимости от имеющихся форматов в документе
        List<Dictionary<string, int>> colValues = new List<Dictionary<string, int>>(); //СПИСОК ЗНАЧЕНИЙ КОЛОНОК
        Dictionary<string, double> totalMeter = new Dictionary<string, double>();
        Thread t_open_file;
        int nextfile= 0;

        public MainWindow()
        {
            InitializeComponent();

            LV.DataContext = table;
            /*
            t_open_file = new Thread(OpenFile_thread);
            t_open_file.Name = "open_file_thread";
            t_open_file.SetApartmentState(ApartmentState.STA);
            t_open_file.IsBackground = true;
            */
        }



        public void OpenFile2()
        {
            int progress_steps = FilePaths.Count - nextfile;    // 2-0=2        //10-6=4
            double progress_value_step = 100 / progress_steps;         //100/2 = 50   //100/4=25
            int progress_current_step = 0;
            Progress_value.Value = progress_value_step * progress_current_step;

            for (int f = nextfile; f < FilePaths.Count; f++)
            {
                var file = FilePaths[f];
                Progress_file_name.Text = file;
                // 1. Открытие файла
                var document = PdfReader.Open(file, PdfDocumentOpenMode.InformationOnly);

                // 2. формируем список страниц с РАЗМЕРами и ФОРМАТами
                int j = 1;
                L_LPages.Add(new List<Page_size>());
                foreach (var page in document.Pages)
                {
                    L_LPages[f].Add(new Page_size(page.Height.Millimeter, page.Width.Millimeter, j));
                    j++;
                }
                progress_current_step++;
                Progress_value.Value = progress_value_step * progress_current_step;
            }

            // 3. формируем СПИСОК КОЛОНОК в зависимости от имеющихся форматов в документе
            for (int i = nextfile; i < L_LPages.Count; i++)
            {
                Progress_file_name.Text = "Формирую колонки";

                foreach (var page in L_LPages[i])
                {
                    if (!colNames.Contains(page.Format))
                        colNames.Add(page.Format);
                }
            }
            colNames.Sort();
            colNames.Reverse();


            //totalMeter = new Dictionary<string, double>();
            //colValues = new List<Dictionary<string, int>>();
            // 4. формируем СПИСОК ЗНАЧЕНИЙ КОЛОНОК
            for(int i = nextfile; i < L_LPages.Count; i++)
            {
                Progress_file_name.Text = "Формирую значения";

                colValues.Add(new Dictionary<string, int>());
                foreach (var page in L_LPages[i])
                {
                    if (!colValues[i].ContainsKey(page.Format))
                        colValues[i].Add(page.Format, 1);
                    else
                        colValues[i][page.Format] += 1;
                    if(page.Format.Equals("Очень большой"))
                    {
                        if (totalMeter.ContainsKey(page.Format)) 
                        {
                            if (page.Width > page.Height)
                                totalMeter[page.Format] += page.Width/1000;
                            else
                                totalMeter[page.Format] += page.Height / 1000;
                        }
                        else
                        {
                            if (page.Width > page.Height)
                                totalMeter[page.Format] = page.Width / 1000;
                            else
                                totalMeter[page.Format] = page.Height / 1000;
                        }
                    }
                }
            }
            UpdateListView2(colNames, colValues);
        }

        private void UpdateListView2(List<string> colNames, List<Dictionary<string, int>> colValues)
        {
            table.Rows.Clear();
            table.Columns.Clear();
            //Progress_file_name.Text = "Формирую таблицу";
            Dispatcher.BeginInvoke(update_prog_status, "Формирую таблицу");


            // Добавляем КОЛОНКИ
            table.Columns.Add("Имя файла", typeof(string));
            foreach (var col in colNames)
            {
                table.Columns.Add(col, typeof(string));
            }

            Dictionary<string, int> total = new Dictionary<string, int>();
            Dictionary<string, int> total3 = new Dictionary<string, int>();

            // Добавляем СТРОКИ
            int r = 0;
            foreach(var list in colValues)
            {
                table.Rows.Add(table.NewRow());
                string fname = FilePaths[r].Substring(FilePaths[r].LastIndexOf('\\') + 1);
                if (fname.Length >50)
                    table.Rows[r]["Имя файла"] = fname.Substring(0,50);
                else
                    table.Rows[r]["Имя файла"] = fname;
                foreach (var word in list)
                {
                    table.Rows[r][word.Key] = word.Value;

                    // попутно заполняем словарь total Итого
                    if (!total.ContainsKey(word.Key))
                        total.Add(word.Key, word.Value);
                    else
                        total[word.Key] += word.Value;
                }
                r++;
            }

            // Заполняем Итого листов по форматам
            // запоминаем сколько метров по каждому формату для вывода в лог
            table.Rows.Add(table.NewRow());
            table.Rows[r]["Имя файла"] = "Итого:";
            foreach (var key in total)
            {
                if (key.Key.Equals("A2"))
                {
                    table.Rows[r][key.Key] = key.Value.ToString() + " - " + Math.Round(key.Value * 0.4, 2) + "м";
                    //totalMeter[key.Key] = key.Value * 0.4;
                }
                else if (key.Key.Equals("A1"))
                {
                    table.Rows[r][key.Key] = key.Value.ToString() + " - " + Math.Round(key.Value * 0.841, 2) + "м";
                    //totalMeter[key.Key] = key.Value * 0.841;
                }
                else if (key.Key.Equals("A0"))
                {
                    table.Rows[r][key.Key] = key.Value.ToString() + " - " + Math.Round(key.Value * 1.19, 2) + "м";
                    //totalMeter[key.Key] = key.Value * 1.19;
                }
                else if (key.Key.Equals("Очень большой"))
                {
                    table.Rows[r][key.Key] = key.Value.ToString() + " - " + Math.Round(totalMeter["Очень большой"], 2) + "м";
                }
                else
                    table.Rows[r][key.Key] = key.Value;
            }
            r++;

            // Заполняем Итого листов по форматам для ТРЕХ экземпляров
            // запоминаем сколько метров по каждому формату для вывода в лог
            int volume = 3;
            table.Rows.Add(table.NewRow());
            table.Rows[r]["Имя файла"] = "Итого на 3 экземпляра:";
            foreach (var key in total)
            {
                if (key.Key.Equals("A2"))
                {
                    table.Rows[r][key.Key] = key.Value * volume + " - " + Math.Round(key.Value* volume * 0.4, 2) + "м";
                    totalMeter[key.Key] = key.Value * volume * 0.4;
                }
                else if (key.Key.Equals("A1"))
                {
                    table.Rows[r][key.Key] = key.Value * volume + " - " + Math.Round(key.Value * volume * 0.841, 2) + "м";
                    totalMeter[key.Key] = key.Value * volume * 0.841;
                }
                else if (key.Key.Equals("A0"))
                {
                    table.Rows[r][key.Key] = key.Value * volume + " - " + Math.Round(key.Value * volume * 1.19, 2) + "м";
                    totalMeter[key.Key] = key.Value * volume * 1.19;
                }
                else if (key.Key.Equals("Очень большой"))
                {
                    table.Rows[r][key.Key] = key.Value * volume + " - " + Math.Round(totalMeter["Очень большой"] * volume, 2) + "м";
                    totalMeter[key.Key] = totalMeter["Очень большой"] * volume ;
                }
                else
                    table.Rows[r][key.Key] = key.Value * volume;
            }

            string log = "";
            double tot = 0.0;
            int i = 0;
            foreach (var key in totalMeter)
            {
                tot += key.Value;
                if (i > 0)
                    log += " + ";
                log += Math.Round(key.Value, 2);
                i++;
            }
            tot = Math.Round(tot, 2);
            log += " = ";

            //txtPlotter.Text = "требуемое количество метров рулона для плоттера: " + log + tot + " метров";
            Dispatcher.BeginInvoke(update_txt_field, log, tot);

            //Progress_file_name.Text = "готово";
            Dispatcher.BeginInvoke(update_prog_status, "Готово");


            GridView gridView = new GridView();
            foreach (DataColumn item in table.Columns)
            {
                GridViewColumn gv_col = new GridViewColumn()
                {
                    Header = item.ColumnName,
                    DisplayMemberBinding = new Binding(item.ColumnName)
                };
                gridView.Columns.Add(gv_col);
            }
            
            LV.View = gridView;
            LV.Items.Refresh();
            

            //Dispatcher.BeginInvoke(update_ui_table, gridView);
            //LV.Dispatcher.Invoke(update_ui_table, gridView);
            //Task task = new Task(() =>LV.View = gridView);
            //Task task2 = new Task(() => LV.Items.Refresh());
            /*
            LV.Dispatcher.BeginInvoke(DispatcherPriority.DataBind, (ThreadStart)
                delegate ()
            {
                LV.View = gridView;
                LV.Items.Refresh();
            });
            */

        }

        void update_progress_bar_value(object obj_value_step, object obj_current_step)
        {
            double progress_value_step = (double)obj_value_step;
            int progress_current_step = (int)obj_current_step;
            Progress_value.Value = progress_value_step * progress_current_step;
        }
        void update_prog_file_name(object obj_file)
        {
            string file = (string)obj_file;
            Progress_file_name.Text = file;
        }

        void update_prog_status(object obj_str)
        {
            string str = (string)obj_str;
            Progress_file_name.Text = str;
        }

        void update_txt_field(object obj_log, object obj_tot)
        {
            string log = (string)obj_log;
            double tot = (double)obj_tot;
            txtPlotter.Text = "требуемое количество метров рулона для плоттера: " + log + tot + " метров";
        }

        public void OpenFile_thread()
        {
            int progress_steps = FilePaths.Count - nextfile;    // 2-0=2        //10-6=4
            double progress_value_step = 100 / progress_steps;         //100/2 = 50   //100/4=25
            int progress_current_step = 0;
            Dispatcher.BeginInvoke(update_progress_bar_value, progress_value_step, progress_current_step);

            for (int f = nextfile; f < FilePaths.Count; f++)
            {
                var file = FilePaths[f];
                Dispatcher.BeginInvoke(update_prog_file_name, file);
                // 1. Открытие файла
                var document = PdfReader.Open(file, PdfDocumentOpenMode.InformationOnly);

                // 2. формируем список страниц с РАЗМЕРами и ФОРМАТами
                int j = 1;
                L_LPages.Add(new List<Page_size>());
                foreach (var page in document.Pages)
                {
                    L_LPages[f].Add(new Page_size(page.Height.Millimeter, page.Width.Millimeter, j));
                    j++;
                }
                progress_current_step++;
                Dispatcher.BeginInvoke(update_progress_bar_value, progress_value_step, progress_current_step);
            }

            // 3. формируем СПИСОК КОЛОНОК в зависимости от имеющихся форматов в документе
            for (int i = nextfile; i < L_LPages.Count; i++)
            {
                Dispatcher.BeginInvoke(update_prog_status, "Формирую колонки");

                foreach (var page in L_LPages[i])
                {
                    if (!colNames.Contains(page.Format))
                        colNames.Add(page.Format);
                }
            }
            colNames.Sort();
            colNames.Reverse();


            //totalMeter = new Dictionary<string, double>();
            //colValues = new List<Dictionary<string, int>>();
            // 4. формируем СПИСОК ЗНАЧЕНИЙ КОЛОНОК
            for (int i = nextfile; i < L_LPages.Count; i++)
            {
                Dispatcher.BeginInvoke(update_prog_status, "Формирую значения");

                colValues.Add(new Dictionary<string, int>());
                foreach (var page in L_LPages[i])
                {
                    if (!colValues[i].ContainsKey(page.Format))
                        colValues[i].Add(page.Format, 1);
                    else
                        colValues[i][page.Format] += 1;
                    if (page.Format.Equals("Очень большой"))
                    {
                        if (totalMeter.ContainsKey(page.Format))
                        {
                            if (page.Width > page.Height)
                                totalMeter[page.Format] += page.Width / 1000;
                            else
                                totalMeter[page.Format] += page.Height / 1000;
                        }
                        else
                        {
                            if (page.Width > page.Height)
                                totalMeter[page.Format] = page.Width / 1000;
                            else
                                totalMeter[page.Format] = page.Height / 1000;
                        }
                    }
                }
            }
            //UpdateListView2(colNames, colValues);
            Dispatcher.BeginInvoke(UpdateListView_thread, colNames, colValues);
        }

        void UpdateListView_thread(object obj_colNames, object obj_colValues)
        {
            List<string> colNames = (List<string>)obj_colNames;
            List<Dictionary<string, int>> colValues = (List<Dictionary<string, int>>)obj_colValues;


            table.Rows.Clear();
            table.Columns.Clear();
            //Progress_file_name.Text = "Формирую таблицу";
            Dispatcher.BeginInvoke(update_prog_status, "Формирую таблицу");


            // Добавляем КОЛОНКИ
            table.Columns.Add("Имя файла", typeof(string));
            foreach (var col in colNames)
            {
                table.Columns.Add(col, typeof(string));
            }

            Dictionary<string, int> total = new Dictionary<string, int>();
            Dictionary<string, int> total3 = new Dictionary<string, int>();

            // Добавляем СТРОКИ
            int r = 0;
            foreach (var list in colValues)
            {
                table.Rows.Add(table.NewRow());
                string fname = FilePaths[r].Substring(FilePaths[r].LastIndexOf('\\') + 1);
                if (fname.Length > 50)
                    table.Rows[r]["Имя файла"] = fname.Substring(0, 50);
                else
                    table.Rows[r]["Имя файла"] = fname;
                foreach (var word in list)
                {
                    table.Rows[r][word.Key] = word.Value;

                    // попутно заполняем словарь total Итого
                    if (!total.ContainsKey(word.Key))
                        total.Add(word.Key, word.Value);
                    else
                        total[word.Key] += word.Value;
                }
                r++;
            }

            // Заполняем Итого листов по форматам
            // запоминаем сколько метров по каждому формату для вывода в лог
            table.Rows.Add(table.NewRow());
            table.Rows[r]["Имя файла"] = "Итого:";
            foreach (var key in total)
            {
                if (key.Key.Equals("A2"))
                {
                    table.Rows[r][key.Key] = key.Value.ToString() + " - " + Math.Round(key.Value * 0.4, 2) + "м";
                    //totalMeter[key.Key] = key.Value * 0.4;
                }
                else if (key.Key.Equals("A1"))
                {
                    table.Rows[r][key.Key] = key.Value.ToString() + " - " + Math.Round(key.Value * 0.841, 2) + "м";
                    //totalMeter[key.Key] = key.Value * 0.841;
                }
                else if (key.Key.Equals("A0"))
                {
                    table.Rows[r][key.Key] = key.Value.ToString() + " - " + Math.Round(key.Value * 1.19, 2) + "м";
                    //totalMeter[key.Key] = key.Value * 1.19;
                }
                else if (key.Key.Equals("Очень большой"))
                {
                    table.Rows[r][key.Key] = key.Value.ToString() + " - " + Math.Round(totalMeter["Очень большой"], 2) + "м";
                }
                else
                    table.Rows[r][key.Key] = key.Value;
            }
            r++;

            // Заполняем Итого листов по форматам для ТРЕХ экземпляров
            // запоминаем сколько метров по каждому формату для вывода в лог
            int volume = 3;
            table.Rows.Add(table.NewRow());
            table.Rows[r]["Имя файла"] = "Итого на 3 экземпляра:";
            foreach (var key in total)
            {
                if (key.Key.Equals("A2"))
                {
                    table.Rows[r][key.Key] = key.Value * volume + " - " + Math.Round(key.Value * volume * 0.4, 2) + "м";
                    totalMeter[key.Key] = key.Value * volume * 0.4;
                }
                else if (key.Key.Equals("A1"))
                {
                    table.Rows[r][key.Key] = key.Value * volume + " - " + Math.Round(key.Value * volume * 0.841, 2) + "м";
                    totalMeter[key.Key] = key.Value * volume * 0.841;
                }
                else if (key.Key.Equals("A0"))
                {
                    table.Rows[r][key.Key] = key.Value * volume + " - " + Math.Round(key.Value * volume * 1.19, 2) + "м";
                    totalMeter[key.Key] = key.Value * volume * 1.19;
                }
                else if (key.Key.Equals("Очень большой"))
                {
                    table.Rows[r][key.Key] = key.Value * volume + " - " + Math.Round(totalMeter["Очень большой"] * volume, 2) + "м";
                    totalMeter[key.Key] = totalMeter["Очень большой"] * volume;
                }
                else
                    table.Rows[r][key.Key] = key.Value * volume;
            }

            string log = "";
            double tot = 0.0;
            int i = 0;
            foreach (var key in totalMeter)
            {
                tot += key.Value;
                if (i > 0)
                    log += " + ";
                log += Math.Round(key.Value, 2);
                i++;
            }
            tot = Math.Round(tot, 2);
            log += " = ";

            //txtPlotter.Text = "требуемое количество метров рулона для плоттера: " + log + tot + " метров";
            Dispatcher.BeginInvoke(update_txt_field, log, tot);
            Dispatcher.BeginInvoke(update_prog_status, "Готово");
            Progress_value.Value = 100.0;

            GridView gridView = new GridView();
            foreach (DataColumn item in table.Columns)
            {
                GridViewColumn gv_col = new GridViewColumn()
                {
                    Header = item.ColumnName,
                    DisplayMemberBinding = new Binding(item.ColumnName)
                };
                gridView.Columns.Add(gv_col);
            }

            LV.View = gridView;
            LV.Items.Refresh();
        }

        public class Page_size
        {
            static Format A4 = new Format("A4", 300.0, 215.0);      // 210x297
            static Format A3 = new Format("A3", 300.0, 425.0);      // 297x420
            static Format A2 = new Format("A2", 431.0, 605.0);      // 420x594
            static Format A1 = new Format("A1", 605.0, 851.0);      // 594x841
            static Format A0 = new Format("A0", 851.0, 1190.0);     // 841x1189
            static Format Other = new Format("Очень большой", 50000.0, 50000.0);      // макс ширина рулона плоттера А0 914мм
            public List<Format> formats = new List<Format>() { A4, A3, A2, A1 ,A0, Other };

            public double Height { get; set; }
            public double Width { get; set; }
            public int Number { get; set; }
            public string Format { get; set; }

            public Page_size(double height, double width, int num) 
            {
                Height = height;
                Width = width;
                Number = num;
                Format = ChoiseFormat(Height, Width, formats, 0);
            }

            private string ChoiseFormat(double H, double W, List<Format> formats, int index)
            {
                if (H <= formats[index].MaxHeight && W <= formats[index].MaxWidth)
                    return formats[index].Name;
                else if(H <= formats[index].MaxWidth && W <= formats[index].MaxHeight)
                    return formats[index].Name;
                else
                    return ChoiseFormat(H, W, formats, index+1);
            }
        }

        public class Format
        {
            public string Name;
            public double MaxHeight;
            public double MaxWidth;
            public Format(string name, double maxHeight, double maxWidth) { 
                Name = name;
                MaxHeight = maxHeight;
                MaxWidth = maxWidth;
            }
        }

        private void btn_AddFile(object sender, RoutedEventArgs e)
        {
            // Выбрать файл и добавить в список
            var fileContent = string.Empty;
            var filePath = string.Empty;

            //using (OpenFileDialog openFileDialog = new OpenFileDialog())
            OpenFileDialog openFileDialog = new OpenFileDialog();
            //openFileDialog.Filter = "pdf files (*.pdf)|*.pdf|All files (*.*)|*.*";
            openFileDialog.Filter = "pdf files (*.pdf)|*.pdf";
            openFileDialog.Multiselect = true;
            openFileDialog.FilterIndex = 1;
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() != null)
            {
                nextfile = FilePaths.Count;
                bool file_added = false;
                foreach (String file in openFileDialog.FileNames)
                {
                    if (!FilePaths.Contains(file))
                    {
                        FilePaths.Add(file);
                        file_added = true;
                    }
                }
                if (file_added)
                {
                    t_open_file = new Thread(OpenFile_thread);
                    t_open_file.Name = "open_file_thread";
                    t_open_file.SetApartmentState(ApartmentState.STA);
                    t_open_file.IsBackground = true;
                    t_open_file.Start();
                }
                //OpenFile2();

            }
        }

        private void btn_Clear(object sender, RoutedEventArgs e)
        {
            FilePaths = new List<string>();
            L_LPages = new List<List<Page_size>>();
            colNames = new List<string>();
            colValues = new List<Dictionary<string, int>>();
            nextfile = 0;
            table.Rows.Clear();
            table.Columns.Clear();
            txtPlotter.Text = "";
        }
    }
}
