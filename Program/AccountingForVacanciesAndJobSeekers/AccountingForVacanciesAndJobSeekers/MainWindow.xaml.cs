using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using DataTable = System.Data.DataTable;
using Window = System.Windows.Window;

namespace AccountingForVacanciesAndJobSeekers
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public int tableIndex;

        //глобальные переменные
        public static int changing = 0;
        public static object element1;
        public static object element2;
        public static object element3;
        public static object element4;
        public static object element5;
        public static object element6;
        public static object element7;
        public static object element8;
        public static object element9;
        public static object element10;

        public System.Windows.Controls.DataGrid pavlyuchkov;

        private List<UIElement> comboBoxElements;

        private List<UIElement> documentElements;

        public MainWindow()
        {
            InitializeComponent();
            //collection for filtration
            comboBoxElements = new List<UIElement>();
            foreach (UIElement element in menuFilter.Items)
            {
                comboBoxElements.Add(element);
            }
            menuFilter.Items.Clear();

            documentElements = new List<UIElement>();
            foreach (UIElement element in documentsMenu.Items)
            {
                documentElements.Add(element);
            }
            documentsMenu.Items.Clear();

            documentsMenu.Visibility = Visibility.Collapsed;
        }

        //-----------------ВЫВОД ТАБЛИЦ-----------------

        private void LoadData(DatabaseConnection dbConnection, string query)
        {
            try
            {
                SqlConnection connection = dbConnection.GetConnection();
                if (dbConnection.OpenConnection())
                {
                    SqlDataAdapter adapter = new SqlDataAdapter(query, connection);
                    System.Data.DataTable dataTable = new System.Data.DataTable();
                    adapter.Fill(dataTable);
                    dataGridForm.ItemsSource = dataTable.DefaultView;
                    dbConnection.CloseConnection();
                }
                else
                {
                    MessageBox.Show("Failed to open connection.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        /*Индексация таблиц:
        Соискатели - 1
        Вакансии - 2
        Заявки - 3
        Октклик - 4
        Приказ - 5
        Образование - 6
        ОпытРаботы - 7*/

        //соискатели - 1
        private void menuTableItems_Click(object sender, RoutedEventArgs e)
        {
            printExselCoic.Visibility = Visibility.Collapsed;
            dataGridFormSecondary.Visibility = Visibility.Collapsed;
            menuFilter.Items.Clear();
            menuFilter.Items.Add(comboBoxElements[0]);
            menuFilter.Items.Add(comboBoxElements[1]);
            menuFilter.Items.Add(comboBoxElements[2]);
            menuFilter.Items.Add(comboBoxElements[3]);

            documentsMenu.Visibility = Visibility.Visible;
            documentsMenu.Items.Clear();
            documentsMenu.Items.Add(documentElements[2]);

            DatabaseConnection dbConnection = new DatabaseConnection();
            string query = @"
        SELECT 
            id_соискателя AS ID, 
            фио AS ФИО, 
            пол AS Пол, 
            FORMAT(дата_рождения, 'dd.MM.yyyy') AS 'Дата рождения', 
            телефон AS Телефон,
            знание_языка AS 'Знание языка' 
        FROM Соискатели";
            LoadData(dbConnection, query);

            // Скроем колонку ID, если это нужно
            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
            tableIndex = 1;
        }

        // Вакансии - 2
        private void menuTableMarks_Click(object sender, RoutedEventArgs e)
        {
            printExselCoic.Visibility = Visibility.Visible;
            dataGridFormSecondary.Visibility = Visibility.Visible;
            menuFilter.Items.Clear();
            menuFilter.Items.Add(comboBoxElements[0]);
            menuFilter.Items.Add(comboBoxElements[1]);
            menuFilter.Items.Add(comboBoxElements[2]);
            menuFilter.Items.Add(comboBoxElements[3]);

            documentsMenu.Visibility = Visibility.Collapsed;
            documentsMenu.Items.Clear();

            DatabaseConnection dbConnection = new DatabaseConnection();
            string query = @"
    SELECT 
        v.id_вакансии AS ID, 
        v.должность AS 'Должность', 
        v.опыт AS 'Опыт', 
        v.зп AS 'Зарплата', 
        v.активность AS 'Активность',
        FORMAT(v.дата_закрытия, 'dd.MM.yyyy') AS 'Дата закрытия',
        vo.наименование AS 'Вид образования'
    FROM Вакансии v
    JOIN ВидыОбразования vo ON v.id_ВидаОбразования = vo.id_ВидаОбразования";

            LoadData(dbConnection, query);
            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
            tableIndex = 2;

            string querySecondary = @"
    SELECT 
        id_соискателя AS ID, 
        фио AS ФИО, 
        пол AS Пол, 
        FORMAT(дата_рождения, 'dd.MM.yyyy') AS 'Дата рождения', 
        телефон AS Телефон,
        знание_языка AS 'Знание языка' 
    FROM Соискатели";

            LoadDataSecondary(dbConnection, querySecondary);
        }


        // Метод загрузки данных во второй DataGrid
        private void LoadDataSecondary(DatabaseConnection dbConnection, string query)
        {
            if (dbConnection.OpenConnection())
            {
                using (SqlCommand command = new SqlCommand(query, dbConnection.GetConnection()))
                {
                    try
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            DataTable dataTable = new DataTable();
                            dataTable.Load(reader);
                            dataGridFormSecondary.ItemsSource = dataTable.DefaultView;
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка при выполнении запроса: " + ex.Message);
                    }
                }
            }
            else
            {
                MessageBox.Show("Ошибка подключения к базе данных.");
            }

            // Проверяем, есть ли колонки в DataGridFormSecondary
            if (dataGridFormSecondary.Columns.Count > 0)
            {
                // Скрываем первую колонку
                dataGridFormSecondary.Columns[0].Visibility = Visibility.Hidden;
            }
        }

        //заявки - 3
        private void menuTableStudent_Click(object sender, RoutedEventArgs e)
        {
            printExselCoic.Visibility = Visibility.Collapsed;
            dataGridFormSecondary.Visibility = Visibility.Collapsed;
            menuFilter.Items.Clear();
            menuFilter.Items.Add(comboBoxElements[0]);
            menuFilter.Items.Add(comboBoxElements[1]);
            menuFilter.Items.Add(comboBoxElements[2]);
            menuFilter.Items.Add(comboBoxElements[3]);

            documentsMenu.Visibility = Visibility.Visible;
            documentsMenu.Items.Clear();
            documentsMenu.Items.Add(documentElements[0]);
            documentsMenu.Items.Add(documentElements[1]);

            DatabaseConnection dbConnection = new DatabaseConnection();
            string query = "SELECT z.id_заявки AS ID, " +
               "v.должность AS 'Должность', " +
               "s.фио AS 'ФИО соискателя', " +
               "z.статус AS 'Статус', " +
               "FORMAT(z.дата, 'dd.MM.yyyy') AS 'Дата' " +
               "FROM Заявки z " +
               "INNER JOIN Соискатели s ON z.id_соискателя = s.id_соискателя " +
               "INNER JOIN Вакансии v ON z.id_вакансии = v.id_вакансии";
            LoadData(dbConnection, query);
            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
            tableIndex = 3;
        }

        //отклик - 4
        private void menuTableSkip_Click(object sender, RoutedEventArgs e)
        {
            printExselCoic.Visibility = Visibility.Collapsed;
            dataGridFormSecondary.Visibility = Visibility.Collapsed;
            menuFilter.Items.Clear();
            menuFilter.Items.Add(comboBoxElements[0]);
            menuFilter.Items.Add(comboBoxElements[1]);
            menuFilter.Items.Add(comboBoxElements[2]);
            menuFilter.Items.Add(comboBoxElements[3]);

            documentsMenu.Visibility = Visibility.Collapsed;
            documentsMenu.Items.Clear();

            try
            {
                DatabaseConnection dbConnection = new DatabaseConnection();
                string query = @"
            SELECT o.id_отклика AS ID, 
                   CONCAT(v.должность,CHAR(13),s.фио) AS 'Заявка', 
                   o.решение AS 'Решение' 
            FROM Отклик o 
            INNER JOIN Заявки z ON o.id_заявки = z.id_заявки 
            INNER JOIN Соискатели s ON z.id_соискателя = s.id_соискателя 
            INNER JOIN Вакансии v ON z.id_вакансии = v.id_вакансии";

                LoadData(dbConnection, query);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }

            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
            tableIndex = 4;
        }

        //приказ - 5
        private void menuTableLisainces_Click(object sender, RoutedEventArgs e)
        {
            printExselCoic.Visibility = Visibility.Collapsed;
            dataGridFormSecondary.Visibility = Visibility.Collapsed;
            menuFilter.Items.Clear();
            menuFilter.Items.Add(comboBoxElements[0]);
            menuFilter.Items.Add(comboBoxElements[1]);
            menuFilter.Items.Add(comboBoxElements[2]);
            menuFilter.Items.Add(comboBoxElements[3]);

            documentsMenu.Visibility = Visibility.Visible;
            documentsMenu.Items.Clear();
            documentsMenu.Items.Add(documentElements[0]);
            documentsMenu.Items.Add(documentElements[1]);

            DatabaseConnection dbConnection = new DatabaseConnection();
            string query = @"
        SELECT 
            p.id_приказа AS ID, 
            CONCAT(v.должность,CHAR(13),s.фио) AS 'Отклик',
            FORMAT(p.дата, 'dd.MM.yyyy') AS 'Дата'
        FROM 
            Приказ p 
        INNER JOIN 
            Отклик o ON p.id_отклика = o.id_отклика 
        INNER JOIN 
            Заявки z ON o.id_заявки = z.id_заявки 
        INNER JOIN 
            Соискатели s ON z.id_соискателя = s.id_соискателя 
        INNER JOIN 
            Вакансии v ON z.id_вакансии = v.id_вакансии";

            LoadData(dbConnection, query);
            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
            tableIndex = 5;
        }


        // Образование - 6
        private void menuTableParents_Click(object sender, RoutedEventArgs e)
        {
            printExselCoic.Visibility = Visibility.Collapsed;
            dataGridFormSecondary.Visibility = Visibility.Collapsed;
            menuFilter.Items.Clear();
            menuFilter.Items.Add(comboBoxElements[0]);
            menuFilter.Items.Add(comboBoxElements[1]);
            menuFilter.Items.Add(comboBoxElements[2]);
            menuFilter.Items.Add(comboBoxElements[3]);

            documentsMenu.Visibility = Visibility.Collapsed;
            documentsMenu.Items.Clear();
            DatabaseConnection dbConnection = new DatabaseConnection();
            string query = "SELECT o.id_образования AS ID, " +
                           "s.фио AS 'ФИО соискателя', " +
                           "o.наименование AS 'Наименование', " +
                           "o.специальность AS 'Специальность', " +
                           "o.номерДиплома AS 'Номер диплома', " +
                           "FORMAT(o.дата_конца, 'dd.MM.yyyy') AS 'Дата конца', " +
                           "v.наименование AS 'Вид образования' " +
                           "FROM Образование o " +
                           "INNER JOIN Соискатели s ON o.id_соискателя = s.id_соискателя " +
                           "INNER JOIN ВидыОбразования v ON o.id_ВидаОбразования = v.id_ВидаОбразования";
            LoadData(dbConnection, query);
            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
            tableIndex = 6;
        }

        // Опыт работы - 7
        private void menuTableEvents_Click(object sender, RoutedEventArgs e)
        {
            printExselCoic.Visibility = Visibility.Collapsed;
            dataGridFormSecondary.Visibility = Visibility.Collapsed;
            menuFilter.Items.Clear();
            menuFilter.Items.Add(comboBoxElements[0]);
            menuFilter.Items.Add(comboBoxElements[1]);
            menuFilter.Items.Add(comboBoxElements[2]);
            menuFilter.Items.Add(comboBoxElements[3]);

            documentsMenu.Visibility = Visibility.Collapsed;
            documentsMenu.Items.Clear();
            DatabaseConnection dbConnection = new DatabaseConnection();
            string query = "SELECT o.id_опыта_работы AS ID, " +
                           "s.фио AS 'ФИО соискателя', " +
                           "o.наименование AS 'Наименование', " +
                           "o.должность AS 'Должность', " +
                           "FORMAT(o.дата_начала, 'dd.MM.yyyy') AS 'Дата начала', " +
                           "FORMAT(o.дата_конца, 'dd.MM.yyyy') AS 'Дата конца' " +
                           "FROM ОпытРаботы o " +
                           "INNER JOIN Соискатели s ON o.id_соискателя = s.id_соискателя";
            LoadData(dbConnection, query);
            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
            tableIndex = 7;
        }

        //Воинский учёт - 8
        private void menuTableVoenka_Click(object sender, RoutedEventArgs e)
        {
            printExselCoic.Visibility = Visibility.Collapsed;
            dataGridFormSecondary.Visibility = Visibility.Collapsed;
            menuFilter.Items.Clear();
            menuFilter.Items.Add(comboBoxElements[0]);
            menuFilter.Items.Add(comboBoxElements[1]);
            menuFilter.Items.Add(comboBoxElements[2]);
            menuFilter.Items.Add(comboBoxElements[3]);

            documentsMenu.Visibility = Visibility.Collapsed;
            documentsMenu.Items.Clear();

            DatabaseConnection dbConnection = new DatabaseConnection();
            string query = @"
        SELECT 
            vu.id_военки AS ID, 
            s.фио AS 'ФИО соискателя', 
            vu.годность AS 'Годность', 
            FORMAT(vu.дата_начала, 'dd.MM.yyyy') AS 'Дата начала', 
            FORMAT(vu.дата_конца, 'dd.MM.yyyy') AS 'Дата конца' 
        FROM 
            ВоинскийУчёт vu 
        INNER JOIN 
            Соискатели s ON vu.id_соискателя = s.id_соискателя";

            LoadData(dbConnection, query);
            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
            tableIndex = 8;
        }

        //Виды образования - 9
        private void menuTableTypeEducational_Click(object sender, RoutedEventArgs e)
        {
            printExselCoic.Visibility = Visibility.Collapsed;
            dataGridFormSecondary.Visibility = Visibility.Collapsed;
            menuFilter.Items.Clear();
            menuFilter.Items.Add(comboBoxElements[0]);
            menuFilter.Items.Add(comboBoxElements[3]);
            documentsMenu.Visibility = Visibility.Collapsed;
            documentsMenu.Items.Clear();

            DatabaseConnection dbConnection = new DatabaseConnection();
            string query = @"
        SELECT 
            id_ВидаОбразования AS 'ID', 
            наименование AS 'Наименование'
        FROM 
            ВидыОбразования";

            LoadData(dbConnection, query);
            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
            tableIndex = 9;
        }
        //-----------------КОНЕЦ ВЫВОД ТАБЛИЦ-----------------

        //-----------------ДОБАВЛЕНИЕ-----------------
        private void menuTableAddedRow_Click(object sender, RoutedEventArgs e)
        {
            switch (tableIndex)
            {
                case 1:
                    // Открыть новую форму Соискатели
                    Соискатели applicantsForm = new Соискатели(this);
                    applicantsForm.Owner = this;
                    applicantsForm.ShowDialog();
                    break;
                case 2:
                    // Открыть новую форму Вакансии
                    Вакансии vacanciesForm = new Вакансии(this);
                    vacanciesForm.Owner = this;
                    vacanciesForm.ShowDialog();
                    break;
                case 3:
                    // Открыть новую форму Заявки
                    Заявки applicationsForm = new Заявки(this);
                    applicationsForm.Owner = this;
                    applicationsForm.ShowDialog();
                    break;
                /*case 4:
                    // Открыть новую форму Отклик
                    Отклик responseForm = new Отклик(this);
                    responseForm.Owner = this;
                    responseForm.ShowDialog();
                    break;
                case 5:
                    // Открыть новую форму Приказ
                    Приказ orderForm = new Приказ(this);
                    orderForm.Owner = this;
                    orderForm.ShowDialog();
                    break;*/
                case 6:
                    // Открыть новую форму Образование
                    Образование educationForm = new Образование(this);
                    educationForm.Owner = this;
                    educationForm.ShowDialog();
                    break;
                case 7:
                    // Открыть новую форму Опыт работы
                    ОпытРаботы experienceForm = new ОпытРаботы(this);
                    experienceForm.Owner = this;
                    experienceForm.ShowDialog();
                    break;
                case 8:
                    // Открыть новую форму воинский учёт
                    ВоинскийУчёт voenForm = new ВоинскийУчёт(this);
                    voenForm.Owner = this;
                    voenForm.ShowDialog();
                    break;
                case 9:
                    MessageBox.Show("Новый тип образования нельзя добавить!");
                    break;
                default:
                    MessageBox.Show("Выберите таблицу!");
                    break;
            }
        }
        //-----------------КОНЕЦ ДОБАВЛЕНИЕ-----------------



        //-----------------ИЗМЕНЕНИЕ-----------------
        private void menuTableChanging_Click(object sender, RoutedEventArgs e)
        {
            DataRowView selectedRow = (DataRowView)dataGridForm.SelectedItem;

            if (selectedRow == null)
            {
                MessageBox.Show("Выберите строку!");
                return;
            }

            switch (tableIndex)
            {
                case 1:
                    element1 = selectedRow["ID"].ToString();
                    element2 = selectedRow["ФИО"].ToString();
                    element3 = selectedRow["Пол"].ToString();
                    element4 = selectedRow["Дата рождения"].ToString();
                    element5 = selectedRow["Телефон"].ToString();
                    element6 = selectedRow["Знание языка"].ToString();
                    changing = 1;
                    Соискатели applicantForm = new Соискатели(this);
                    applicantForm.Owner = this;
                    applicantForm.ShowDialog();
                    break;
                case 2:
                    element1 = selectedRow["ID"].ToString();
                    element2 = selectedRow["Должность"].ToString();
                    element3 = selectedRow["Опыт"].ToString();
                    element4 = selectedRow["Зарплата"].ToString();
                    element5 = selectedRow["Активность"].ToString();
                    element6 = selectedRow["Вид образования"].ToString();
                    changing = 1;
                    Вакансии vacanciesForm = new Вакансии(this);
                    vacanciesForm.Owner = this;
                    vacanciesForm.ShowDialog();
                    break;
                case 3:
                    element1 = selectedRow["ID"].ToString();
                    element2 = selectedRow["Должность"].ToString();
                    element3 = selectedRow["ФИО соискателя"].ToString();
                    element4 = selectedRow["Статус"].ToString();
                    changing = 1;
                    Заявки applicationsForm = new Заявки(this);
                    applicationsForm.Owner = this;
                    applicationsForm.ShowDialog();
                    break;
                /*case 4:
                    element1 = selectedRow["ID"].ToString();
                    element2 = selectedRow["Заявка"].ToString();
                    element3 = selectedRow["Решение"].ToString();
                    changing = 1;
                    Отклик responseForm = new Отклик(this);
                    responseForm.Owner = this;
                    responseForm.ShowDialog();
                    break;
                case 5:
                    element1 = selectedRow["ID"].ToString();
                    element2 = selectedRow["Отклик"].ToString();
                    element3 = selectedRow["Дата"].ToString();
                    changing = 1;
                    Приказ orderForm = new Приказ(this);
                    orderForm.Owner = this;
                    orderForm.ShowDialog();
                    break;*/
                case 6:
                    element1 = selectedRow["ID"].ToString();
                    element2 = selectedRow["ФИО соискателя"].ToString();
                    element3 = selectedRow["Наименование"].ToString();
                    element4 = selectedRow["Специальность"].ToString();
                    element5 = selectedRow["Номер диплома"].ToString();
                    element6 = selectedRow["Дата конца"].ToString();
                    element7 = selectedRow["Вид образования"].ToString();
                    changing = 1;
                    Образование educationForm = new Образование(this);
                    educationForm.Owner = this;
                    educationForm.ShowDialog();
                    break;
                case 7:
                    element1 = selectedRow["ID"].ToString();
                    element2 = selectedRow["ФИО соискателя"].ToString();
                    element3 = selectedRow["Наименование"].ToString();
                    element4 = selectedRow["Должность"].ToString();
                    element5 = selectedRow["Дата начала"].ToString();
                    element6 = selectedRow["Дата конца"].ToString();
                    changing = 1;
                    ОпытРаботы experienceForm = new ОпытРаботы(this);
                    experienceForm.Owner = this;
                    experienceForm.ShowDialog();
                    break;
                case 8:
                    element1 = selectedRow["ID"].ToString();
                    element2 = selectedRow["ФИО соискателя"].ToString();
                    element3 = selectedRow["Годность"].ToString();
                    element4 = selectedRow["Дата начала"].ToString();
                    element5 = selectedRow["Дата конца"].ToString();
                    changing = 1;
                    ВоинскийУчёт voenForm = new ВоинскийУчёт(this);
                    voenForm.Owner = this;
                    voenForm.ShowDialog();
                    break;
                case 9:
                    MessageBox.Show("Тип образования нельзя изменить!");
                    break;
                default:
                    MessageBox.Show("Выберите таблицу!");
                    break;
            }
        }

        //-----------------КОНЕЦ ИЗМЕНЕНИЕ-----------------


        //-----------------ОБНОВЛЕНИЕ-----------------
        public void menuTableRefresh_Click(object sender, RoutedEventArgs e)
        {
            Refresh(sender, e);
        }

        public void Refresh(object sender, RoutedEventArgs e)
        {
            switch (tableIndex)
            {
                case 1:
                    menuTableItems_Click(sender, e);
                    break;
                case 2:
                    menuTableMarks_Click(sender, e);
                    break;
                case 3:
                    menuTableStudent_Click(sender, e);
                    break;
                case 4:
                    menuTableSkip_Click(sender, e);
                    break;
                case 5:
                    menuTableLisainces_Click(sender, e);
                    break;
                case 6:
                    menuTableParents_Click(sender, e);
                    break;
                case 7:
                    menuTableEvents_Click(sender, e);
                    break;
                case 8:
                    menuTableVoenka_Click(sender, e);
                    break;
                case 9:
                    menuTableTypeEducational_Click(sender, e);
                    break;
                default:
                    MessageBox.Show("Выберите таблицу!");
                    break;
            }
        }
        //-----------------КОНЕЦ ОБНОВЛЕНИЕ-----------------

        //-----------------УДАЛЕНИЕ-----------------
        private void menuTableDelete_Click(object sender, RoutedEventArgs e)
        {
            // Получение выбранной строки из DataGrid
            DataRowView selectedRow = (DataRowView)dataGridForm.SelectedItem;

            // Проверка, что строка действительно выбрана
            if (selectedRow != null)
            {
                try
                {
                    // Получение идентификатора из первой колонки (предполагается, что идентификатор находится в первой колонке)
                    int idToDelete = Convert.ToInt32(selectedRow[0]);

                    // Выполнение операции удаления в базе данных в зависимости от выбранной таблицы
                    DatabaseConnection dbConnection = new DatabaseConnection();
                    SqlConnection connection = dbConnection.GetConnection();
                    if (dbConnection.OpenConnection())
                    {
                        SqlCommand command = null;
                        switch (tableIndex)
                        {
                            case 1:
                                command = new SqlCommand("DELETE FROM Соискатели WHERE id_соискателя = @ID", connection);
                                break;
                            case 2:
                                command = new SqlCommand("DELETE FROM Вакансии WHERE id_вакансии = @ID", connection);
                                break;
                            case 3:
                                command = new SqlCommand("DELETE FROM Заявки WHERE id_заявки = @ID", connection);
                                break;
                            case 4:
                                command = new SqlCommand("DELETE FROM Отклик WHERE id_отклика = @ID", connection);
                                break;
                            case 5:
                                command = new SqlCommand("DELETE FROM Приказ WHERE id_приказа = @ID", connection);
                                break;
                            case 6:
                                command = new SqlCommand("DELETE FROM Образование WHERE id_образования = @ID", connection);
                                break;
                            case 7:
                                command = new SqlCommand("DELETE FROM ОпытРаботы WHERE id_опыта_работы = @ID", connection);
                                break;
                            case 8:
                                command = new SqlCommand("DELETE FROM ВоинскийУчёт WHERE id_военки = @ID", connection);
                                break;
                            case 9:
                                MessageBox.Show("Тип образования нельзя удалить!");
                                break;
                            default:
                                MessageBox.Show("Выберите таблицу!");
                                return; // Прекращаем выполнение метода, так как нет команды для удаления
                        }

                        // Установка параметра и выполнение команды удаления
                        if (command != null)
                        {
                            command.Parameters.AddWithValue("@ID", idToDelete);
                            command.ExecuteNonQuery();
                        }

                        // Обновление DataGrid после удаления
                        Refresh(sender, e);

                        dbConnection.CloseConnection();
                    }
                    else
                    {
                        MessageBox.Show("Failed to open connection.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка при удалении записи: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Выберите строку для удаления!");
            }
        }
        //-----------------КОНЕЦ УДАЛЕНИЕ-----------------


        //-----------------ВЫВОД В EXCEL-----------------
        private void printExsel_Click(object sender, RoutedEventArgs e)
        {
            if (tableIndex == 0)
            {
                MessageBox.Show("Выберите таблицу!");
                return;
            }

            string nameTable = string.Empty;

            // Устанавливаем имя таблицы в зависимости от выбора пользователя
            switch (tableIndex)
            {
                case 1:
                    nameTable = "Соискатели";
                    break;
                case 2:
                    nameTable = "Вакансии";
                    break;
                case 3:
                    nameTable = "Заявки";
                    break;
                case 4:
                    nameTable = "Отклик";
                    break;
                case 5:
                    nameTable = "Приказ";
                    break;
                case 6:
                    nameTable = "Образование";
                    break;
                case 7:
                    nameTable = "Опыт работы";
                    break;
            }

            // Создание объекта SaveFileDialog
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Файлы Excel (*.xlsx)|*.xlsx";
            saveFileDialog.Title = "Сохранить как Excel";
            saveFileDialog.DefaultExt = "xlsx";

            if (saveFileDialog.ShowDialog() == true)
            {
                // Получение выбранного пользователем пути и имени файла
                string filePath = saveFileDialog.FileName;

                // Создание нового объекта приложения Excel
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Visible = false; // Скрываем Excel

                // Создание новой книги Excel
                Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Add(Type.Missing);

                // Создание нового листа Excel
                Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets[1];

                // Заполнение листа данными из вашего DataGrid
                for (int i = 0; i < dataGridForm.Items.Count; i++)
                {
                    var dataGridRow = (DataGridRow)dataGridForm.ItemContainerGenerator.ContainerFromIndex(i);
                    if (dataGridRow != null)
                    {
                        for (int j = 0; j < dataGridForm.Columns.Count; j++)
                        {
                            var content = dataGridForm.Columns[j].GetCellContent(dataGridRow);
                            if (content is TextBlock)
                            {
                                var text = (content as TextBlock).Text;
                                excelSheet.Cells[i + 2, j + 1] = text; // Начинаем с второй строки

                                // Если это столбец с телефонным номером, установить формат ячейки в текстовый
                                if (dataGridForm.Columns[j].Header.ToString() == "Телефон")
                                {
                                    excelSheet.Cells[i + 2, j + 1].NumberFormat = "@";
                                }
                            }
                        }
                    }
                }


                // Удаление столбца A
                Microsoft.Office.Interop.Excel.Range columnA = (Microsoft.Office.Interop.Excel.Range)excelSheet.Columns["A"];
                columnA.Delete();

                // Объединение ячеек в первой строке
                Microsoft.Office.Interop.Excel.Range headerRange = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[1, dataGridForm.Columns.Count - 1]];
                headerRange.Merge();

                // Установка текста в объединенной ячейке
                excelSheet.Cells[1, 1] = nameTable;

                // Выравнивание текста по центру и установка жирного шрифта для первой строки
                headerRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                headerRange.Font.Bold = true;

                // Добавление обводки для всей таблицы Excel
                Microsoft.Office.Interop.Excel.Range tableRange = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[dataGridForm.Items.Count + 1, dataGridForm.Columns.Count - 1]];
                tableRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                tableRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                // Выравнивание ширины столбцов
                excelSheet.UsedRange.Columns.AutoFit();

                // Сохранение книги Excel по выбранному пути
                excelBook.SaveAs(filePath);

                // Закрытие книги и приложения Excel
                excelBook.Close();
                excelApp.Quit();

                // Освобождение ресурсов COM
                Marshal.ReleaseComObject(excelSheet);
                Marshal.ReleaseComObject(excelBook);
                Marshal.ReleaseComObject(excelApp);
            }
        }
        //-----------------КОНЕЦ ВЫВОД В EXCEL-----------------


        //-----------------ФИЛЬТРАЦИЯ-----------------
        private void buttonFilter_Click(object sender, RoutedEventArgs e)
        {
            switch (tableIndex)
            {
                case 1: // Соискатели
                    try
                    {
                        string filterText = textBoxFilter.Text.Trim(); // Получаем текст фильтрации
                        DateTime? startDate = datePickerFilterFirstDate.SelectedDate;
                        DateTime? endDate = datePickerFilterLastDate.SelectedDate;

                        // Формируем SQL-запрос для фильтрации и выборки данных
                        string sqlQuery = @"
            SELECT 
                id_соискателя AS ID, 
                фио AS 'ФИО', 
                пол AS 'Пол', 
                FORMAT(дата_рождения, 'dd.MM.yyyy') AS 'Дата рождения', 
                телефон AS 'Телефон',
                знание_языка AS 'Знание языка'
            FROM Соискатели 
            WHERE 
                (
                    фио LIKE @FilterText OR 
                    пол LIKE @FilterText OR 
                    FORMAT(дата_рождения, 'dd.MM.yyyy') LIKE @FilterText OR 
                    телефон LIKE @FilterText OR
                    знание_языка LIKE @FilterText
                ) AND 
                (@StartDate IS NULL OR дата_рождения >= @StartDate) AND 
                (@EndDate IS NULL OR дата_рождения <= @EndDate)";

                        // Создаем подключение к базе данных
                        using (DatabaseConnection dbConnection = new DatabaseConnection())
                        {
                            if (dbConnection.OpenConnection())
                            {
                                using (SqlCommand command = new SqlCommand(sqlQuery, dbConnection.GetConnection()))
                                {
                                    command.Parameters.AddWithValue("@FilterText", "%" + filterText + "%");
                                    command.Parameters.AddWithValue("@StartDate", startDate ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@EndDate", endDate ?? (object)DBNull.Value);

                                    try
                                    {
                                        using (SqlDataReader reader = command.ExecuteReader())
                                        {
                                            DataTable dataTable = new DataTable();
                                            dataTable.Load(reader);
                                            dataGridForm.ItemsSource = dataTable.DefaultView;
                                            // Скрываем колонку ID, если это нужно
                                            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
                                            // Здесь вы можете использовать dataTable для отображения результатов
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show("Ошибка при выполнении запроса: " + ex.Message);
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("Ошибка подключения к базе данных.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка: " + ex.Message);
                    }
                    break;
                case 2: // Вакансии
                    try
                    {
                        string filterText = textBoxFilter.Text.Trim(); // Получаем текст фильтрации
                        DateTime? startDate = datePickerFilterFirstDate.SelectedDate;
                        DateTime? endDate = datePickerFilterLastDate.SelectedDate;

                        // Формируем SQL-запрос для фильтрации и выборки данных
                        string sqlQuery = @"
        SELECT 
            id_вакансии AS ID, 
            должность AS 'Должность', 
            опыт AS 'Опыт', 
            зп AS 'Зарплата', 
            активность AS 'Активность',
            FORMAT(дата_закрытия, 'dd.MM.yyyy') AS 'Дата закрытия'
        FROM Вакансии 
        WHERE 
            (должность LIKE @FilterText OR 
            опыт LIKE @FilterText OR 
            зп LIKE @FilterText OR 
            активность LIKE @FilterText) 
            AND (@StartDate IS NULL OR дата_закрытия >= @StartDate) 
            AND (@EndDate IS NULL OR дата_закрытия <= @EndDate)";

                        // Создаем подключение к базе данных
                        using (DatabaseConnection dbConnection = new DatabaseConnection())
                        {
                            if (dbConnection.OpenConnection())
                            {
                                using (SqlCommand command = new SqlCommand(sqlQuery, dbConnection.GetConnection()))
                                {
                                    command.Parameters.AddWithValue("@FilterText", "%" + filterText + "%");
                                    command.Parameters.AddWithValue("@StartDate", startDate ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@EndDate", endDate ?? (object)DBNull.Value);

                                    try
                                    {
                                        using (SqlDataReader reader = command.ExecuteReader())
                                        {
                                            DataTable dataTable = new DataTable();
                                            dataTable.Load(reader);
                                            dataGridForm.ItemsSource = dataTable.DefaultView;
                                            // Скрываем колонку ID, если это нужно
                                            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show("Ошибка при выполнении запроса: " + ex.Message);
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("Ошибка подключения к базе данных.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка: " + ex.Message);
                    }
                    break;
                case 3: // Заявки
                    try
                    {
                        string filterText = textBoxFilter.Text.Trim(); // Получаем текст фильтрации
                        DateTime? startDate = datePickerFilterFirstDate.SelectedDate;
                        DateTime? endDate = datePickerFilterLastDate.SelectedDate;

                        // Формируем SQL-запрос для фильтрации и выборки данных
                        string sqlQuery = @"
            SELECT 
                z.id_заявки AS ID, 
                v.должность AS 'Должность', 
                s.фио AS 'ФИО соискателя', 
                z.статус AS 'Статус', 
                FORMAT(z.дата, 'dd.MM.yyyy') AS 'Дата' 
            FROM Заявки z 
            INNER JOIN Соискатели s ON z.id_соискателя = s.id_соискателя 
            INNER JOIN Вакансии v ON z.id_вакансии = v.id_вакансии 
            WHERE 
                (v.должность LIKE @FilterText OR 
                s.фио LIKE @FilterText OR 
                z.статус LIKE @FilterText) 
                AND (@StartDate IS NULL OR z.дата >= @StartDate) 
                AND (@EndDate IS NULL OR z.дата <= @EndDate)";

                        // Создаем подключение к базе данных
                        using (DatabaseConnection dbConnection = new DatabaseConnection())
                        {
                            if (dbConnection.OpenConnection())
                            {
                                using (SqlCommand command = new SqlCommand(sqlQuery, dbConnection.GetConnection()))
                                {
                                    command.Parameters.AddWithValue("@FilterText", "%" + filterText + "%");
                                    command.Parameters.AddWithValue("@StartDate", startDate ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@EndDate", endDate ?? (object)DBNull.Value);

                                    try
                                    {
                                        using (SqlDataReader reader = command.ExecuteReader())
                                        {
                                            DataTable dataTable = new DataTable();
                                            dataTable.Load(reader);
                                            dataGridForm.ItemsSource = dataTable.DefaultView;
                                            // Скрываем колонку ID, если это нужно
                                            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show("Ошибка при выполнении запроса: " + ex.Message);
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("Ошибка подключения к базе данных.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка: " + ex.Message);
                    }
                    break;
                case 4: // Отклик
                    try
                    {
                        string filterText = textBoxFilter.Text.Trim(); // Получаем текст фильтрации

                        // Формируем SQL-запрос для фильтрации и выборки данных
                        string sqlQuery = @"
            SELECT 
                o.id_отклика AS ID, 
                CONCAT(v.должность, CHAR(13), s.фио) AS 'Заявка', 
                o.решение AS 'Решение' 
            FROM Отклик o 
            INNER JOIN Заявки z ON o.id_заявки = z.id_заявки 
            INNER JOIN Соискатели s ON z.id_соискателя = s.id_соискателя 
            INNER JOIN Вакансии v ON z.id_вакансии = v.id_вакансии 
            WHERE 
                (v.должность LIKE @FilterText OR 
                 s.фио LIKE @FilterText OR 
                 o.решение LIKE @FilterText)";

                        // Создаем подключение к базе данных
                        using (DatabaseConnection dbConnection = new DatabaseConnection())
                        {
                            if (dbConnection.OpenConnection())
                            {
                                using (SqlCommand command = new SqlCommand(sqlQuery, dbConnection.GetConnection()))
                                {
                                    command.Parameters.AddWithValue("@FilterText", "%" + filterText + "%");

                                    try
                                    {
                                        using (SqlDataReader reader = command.ExecuteReader())
                                        {
                                            DataTable dataTable = new DataTable();
                                            dataTable.Load(reader);
                                            dataGridForm.ItemsSource = dataTable.DefaultView;
                                            // Скрываем колонку ID, если это нужно
                                            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show("Ошибка при выполнении запроса: " + ex.Message);
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("Ошибка подключения к базе данных.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка: " + ex.Message);
                    }
                    break;
                case 5: // Приказ
                    try
                    {
                        string filterText = textBoxFilter.Text.Trim(); // Получаем текст фильтрации
                        DateTime? startDate = datePickerFilterFirstDate.SelectedDate;
                        DateTime? endDate = datePickerFilterLastDate.SelectedDate;

                        // Формируем SQL-запрос для фильтрации и выборки данных
                        string sqlQuery = @"
            SELECT 
                p.id_приказа AS ID, 
                CONCAT(v.должность, CHAR(13), s.фио) AS 'Отклик',
                FORMAT(p.дата, 'dd.MM.yyyy') AS 'Дата'
            FROM 
                Приказ p 
                INNER JOIN Отклик o ON p.id_отклика = o.id_отклика 
                INNER JOIN Заявки z ON o.id_заявки = z.id_заявки 
                INNER JOIN Соискатели s ON z.id_соискателя = s.id_соискателя 
                INNER JOIN Вакансии v ON z.id_вакансии = v.id_вакансии 
            WHERE 
                (v.должность LIKE @FilterText OR 
                s.фио LIKE @FilterText OR 
                o.решение LIKE @FilterText OR 
                FORMAT(p.дата, 'dd.MM.yyyy') LIKE @FilterText) 
                AND (@StartDate IS NULL OR p.дата >= @StartDate) 
                AND (@EndDate IS NULL OR p.дата <= @EndDate)";

                        // Создаем подключение к базе данных
                        using (DatabaseConnection dbConnection = new DatabaseConnection())
                        {
                            if (dbConnection.OpenConnection())
                            {
                                using (SqlCommand command = new SqlCommand(sqlQuery, dbConnection.GetConnection()))
                                {
                                    command.Parameters.AddWithValue("@FilterText", "%" + filterText + "%");
                                    command.Parameters.AddWithValue("@StartDate", startDate ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@EndDate", endDate ?? (object)DBNull.Value);

                                    try
                                    {
                                        using (SqlDataReader reader = command.ExecuteReader())
                                        {
                                            DataTable dataTable = new DataTable();
                                            dataTable.Load(reader);
                                            dataGridForm.ItemsSource = dataTable.DefaultView;
                                            // Скрываем колонку ID, если это нужно
                                            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show("Ошибка при выполнении запроса: " + ex.Message);
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("Ошибка подключения к базе данных.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка: " + ex.Message);
                    }
                    break;
                case 6: // Образование
                    try
                    {
                        string filterText = textBoxFilter.Text.Trim(); // Получаем текст фильтрации
                        DateTime? startDate = datePickerFilterFirstDate.SelectedDate;
                        DateTime? endDate = datePickerFilterLastDate.SelectedDate;

                        // Формируем SQL-запрос для фильтрации и выборки данных
                        string sqlQuery = @"
            SELECT 
                o.id_образования AS ID, 
                s.фио AS 'ФИО соискателя', 
                o.наименование AS 'Наименование', 
                o.специальность AS 'Специальность', 
                o.номерДиплома AS 'Номер диплома', 
                FORMAT(o.дата_конца, 'dd.MM.yyyy') AS 'Дата конца', 
                v.наименование AS 'Вид образования' 
            FROM 
                Образование o 
            INNER JOIN 
                Соискатели s ON o.id_соискателя = s.id_соискателя 
            INNER JOIN 
                ВидыОбразования v ON o.id_ВидаОбразования = v.id_ВидаОбразования
            WHERE 
                (s.фио LIKE @FilterText OR 
                o.наименование LIKE @FilterText OR 
                o.специальность LIKE @FilterText OR 
                o.номерДиплома LIKE @FilterText) 
                AND (@StartDate IS NULL OR o.дата_конца >= @StartDate) 
                AND (@EndDate IS NULL OR o.дата_конца <= @EndDate)";

                        // Создаем подключение к базе данных
                        using (DatabaseConnection dbConnection = new DatabaseConnection())
                        {
                            if (dbConnection.OpenConnection())
                            {
                                using (SqlCommand command = new SqlCommand(sqlQuery, dbConnection.GetConnection()))
                                {
                                    command.Parameters.AddWithValue("@FilterText", "%" + filterText + "%");
                                    command.Parameters.AddWithValue("@StartDate", startDate ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@EndDate", endDate ?? (object)DBNull.Value);

                                    try
                                    {
                                        using (SqlDataReader reader = command.ExecuteReader())
                                        {
                                            DataTable dataTable = new DataTable();
                                            dataTable.Load(reader);
                                            dataGridForm.ItemsSource = dataTable.DefaultView;
                                            // Скрываем колонку ID, если это нужно
                                            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show("Ошибка при выполнении запроса: " + ex.Message);
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("Ошибка подключения к базе данных.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка: " + ex.Message);
                    }
                    break;
                case 7: // Опыт работы
                    try
                    {
                        string filterText = textBoxFilter.Text.Trim(); // Получаем текст фильтрации
                        DateTime? startDate = datePickerFilterFirstDate.SelectedDate;
                        DateTime? endDate = datePickerFilterLastDate.SelectedDate;

                        // Формируем SQL-запрос для фильтрации и выборки данных
                        string sqlQuery = @"
            SELECT 
                ОпытРаботы.id_опыта_работы AS ID, 
                Соискатели.фио AS 'ФИО соискателя', 
                ОпытРаботы.наименование AS 'Наименование', 
                ОпытРаботы.должность AS 'Должность', 
                FORMAT(ОпытРаботы.дата_начала, 'dd.MM.yyyy') AS 'Дата начала', 
                FORMAT(ОпытРаботы.дата_конца, 'dd.MM.yyyy') AS 'Дата конца'
            FROM 
                ОпытРаботы
            INNER JOIN 
                Соискатели ON ОпытРаботы.id_соискателя = Соискатели.id_соискателя
            WHERE 
                (Соискатели.фио LIKE @FilterText OR 
                ОпытРаботы.наименование LIKE @FilterText OR 
                ОпытРаботы.должность LIKE @FilterText OR 
                FORMAT(ОпытРаботы.дата_начала, 'dd.MM.yyyy') LIKE @FilterText OR 
                FORMAT(ОпытРаботы.дата_конца, 'dd.MM.yyyy') LIKE @FilterText) 
                AND (@StartDate IS NULL OR ОпытРаботы.дата_начала >= @StartDate) 
                AND (@EndDate IS NULL OR ОпытРаботы.дата_конца <= @EndDate)";

                        // Создаем подключение к базе данных
                        using (DatabaseConnection dbConnection = new DatabaseConnection())
                        {
                            if (dbConnection.OpenConnection())
                            {
                                using (SqlCommand command = new SqlCommand(sqlQuery, dbConnection.GetConnection()))
                                {
                                    command.Parameters.AddWithValue("@FilterText", "%" + filterText + "%");
                                    command.Parameters.AddWithValue("@StartDate", startDate ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@EndDate", endDate ?? (object)DBNull.Value);

                                    try
                                    {
                                        using (SqlDataReader reader = command.ExecuteReader())
                                        {
                                            DataTable dataTable = new DataTable();
                                            dataTable.Load(reader);
                                            dataGridForm.ItemsSource = dataTable.DefaultView;
                                            // Скрываем колонку ID, если это нужно
                                            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show("Ошибка при выполнении запроса: " + ex.Message);
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("Ошибка подключения к базе данных.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка: " + ex.Message);
                    }
                    break;
                case 8: // Воинский учёт
                    try
                    {
                        string filterText = textBoxFilter.Text.Trim(); // Получаем текст фильтрации
                        DateTime? startDate = datePickerFilterFirstDate.SelectedDate;
                        DateTime? endDate = datePickerFilterLastDate.SelectedDate;

                        // Формируем SQL-запрос для фильтрации и выборки данных
                        string sqlQuery = @"
            SELECT 
                vu.id_военки AS ID, 
                s.фио AS 'ФИО соискателя', 
                vu.годность AS 'Годность', 
                FORMAT(vu.дата_начала, 'dd.MM.yyyy') AS 'Дата начала', 
                FORMAT(vu.дата_конца, 'dd.MM.yyyy') AS 'Дата конца' 
            FROM 
                ВоинскийУчёт vu 
            INNER JOIN 
                Соискатели s ON vu.id_соискателя = s.id_соискателя
            WHERE 
                (s.фио LIKE @FilterText OR 
                 vu.годность LIKE @FilterText OR 
                 FORMAT(vu.дата_начала, 'dd.MM.yyyy') LIKE @FilterText OR 
                 FORMAT(vu.дата_конца, 'dd.MM.yyyy') LIKE @FilterText)
                AND (@StartDate IS NULL OR vu.дата_начала >= @StartDate)
                AND (@EndDate IS NULL OR vu.дата_конца <= @EndDate)";

                        // Создаем подключение к базе данных
                        using (DatabaseConnection dbConnection = new DatabaseConnection())
                        {
                            if (dbConnection.OpenConnection())
                            {
                                using (SqlCommand command = new SqlCommand(sqlQuery, dbConnection.GetConnection()))
                                {
                                    command.Parameters.AddWithValue("@FilterText", "%" + filterText + "%");
                                    command.Parameters.AddWithValue("@StartDate", (object)startDate ?? DBNull.Value);
                                    command.Parameters.AddWithValue("@EndDate", (object)endDate ?? DBNull.Value);

                                    try
                                    {
                                        using (SqlDataReader reader = command.ExecuteReader())
                                        {
                                            DataTable dataTable = new DataTable();
                                            dataTable.Load(reader);
                                            dataGridForm.ItemsSource = dataTable.DefaultView;
                                            // Скрываем колонку ID, если это нужно
                                            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show("Ошибка при выполнении запроса: " + ex.Message);
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("Ошибка подключения к базе данных.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка: " + ex.Message);
                    }
                    break;
                case 9: // Виды образования
                    try
                    {
                        string filterText = textBoxFilter.Text.Trim(); // Получаем текст фильтрации

                        // Формируем SQL-запрос для фильтрации и выборки данных
                        string sqlQuery = @"
            SELECT 
                id_ВидаОбразования AS ID, 
                наименование AS 'Наименование' 
            FROM ВидыОбразования 
            WHERE наименование LIKE @FilterText";

                        // Создаем подключение к базе данных
                        using (DatabaseConnection dbConnection = new DatabaseConnection())
                        {
                            if (dbConnection.OpenConnection())
                            {
                                using (SqlCommand command = new SqlCommand(sqlQuery, dbConnection.GetConnection()))
                                {
                                    command.Parameters.AddWithValue("@FilterText", "%" + filterText + "%");

                                    try
                                    {
                                        using (SqlDataReader reader = command.ExecuteReader())
                                        {
                                            DataTable dataTable = new DataTable();
                                            dataTable.Load(reader);
                                            dataGridForm.ItemsSource = dataTable.DefaultView;
                                            // Скрываем колонку ID, если это нужно
                                            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show("Ошибка при выполнении запроса: " + ex.Message);
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("Ошибка подключения к базе данных.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка: " + ex.Message);
                    }
                    break;
                default:
                    MessageBox.Show("Выберите таблицу!");
                    break;
            }
        }
        //-----------------КОНЕЦ ФИЛЬТРАЦИЯ-----------------


        //-----------------ПОИСК-----------------
        private void buttonSearch_Click(object sender, RoutedEventArgs e)
        {
            string searchText = txtSearch.Text;

            if(searchText == "")
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(searchText))
            {
                ClearSearchHighlighting(dataGridForm);
                ClearSearchHighlighting(dataGridFormSecondary);
                return;
            }

            HighlightSearchText(dataGridForm, searchText);
            HighlightSearchText(dataGridFormSecondary, searchText);
        }

        private void HighlightSearchText(DataGrid dataGrid, string searchText)
        {
            foreach (DataGridRow row in GetDataGridRows(dataGrid))
            {
                foreach (DataGridColumn column in dataGrid.Columns)
                {
                    if (column is DataGridTextColumn)
                    {
                        var cell = GetCell(row, column, dataGrid);
                        if (cell != null)
                        {
                            TextBlock textBlock = cell.Content as TextBlock;
                            if (textBlock != null)
                            {
                                string cellContent = textBlock.Text;
                                if (!string.IsNullOrEmpty(cellContent))
                                {
                                    if (cellContent.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        int index = cellContent.IndexOf(searchText, StringComparison.OrdinalIgnoreCase);
                                        string preMatch = cellContent.Substring(0, index);
                                        string match = cellContent.Substring(index, searchText.Length);
                                        string postMatch = cellContent.Substring(index + searchText.Length);

                                        textBlock.Inlines.Clear();
                                        textBlock.Inlines.Add(new Run(preMatch));
                                        Run matchRun = new Run(match);
                                        matchRun.Background = Brushes.Yellow;
                                        textBlock.Inlines.Add(matchRun);
                                        textBlock.Inlines.Add(new Run(postMatch));
                                    }
                                    else
                                    {
                                        textBlock.Inlines.Clear();
                                        textBlock.Inlines.Add(new Run(cellContent));
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private void ClearSearchHighlighting(DataGrid dataGrid)
        {
            foreach (DataGridRow row in GetDataGridRows(dataGrid))
            {
                foreach (DataGridColumn column in dataGrid.Columns)
                {
                    if (column is DataGridTextColumn)
                    {
                        var cell = GetCell(row, column, dataGrid);
                        if (cell != null)
                        {
                            TextBlock textBlock = cell.Content as TextBlock;
                            if (textBlock != null)
                            {
                                textBlock.Inlines.Clear();
                                textBlock.Inlines.Add(new Run(textBlock.Text));
                            }
                        }
                    }
                }
            }
        }

        private System.Windows.Controls.DataGridCell GetCell(DataGridRow row, DataGridColumn column, DataGrid dataGrid)
        {
            if (column != null)
            {
                DataGridCellsPresenter presenter = GetVisualChild<DataGridCellsPresenter>(row);
                if (presenter == null)
                    return null;

                int columnIndex = dataGrid.Columns.IndexOf(column);
                if (columnIndex > -1)
                {
                    System.Windows.Controls.DataGridCell cell = (System.Windows.Controls.DataGridCell)presenter.ItemContainerGenerator.ContainerFromIndex(columnIndex);
                    return cell;
                }
            }
            return null;
        }

        private List<DataGridRow> GetDataGridRows(System.Windows.Controls.DataGrid grid)
        {
            List<DataGridRow> rows = new List<DataGridRow>();
            for (int i = 0; i < grid.Items.Count; i++)
            {
                DataGridRow row = (DataGridRow)grid.ItemContainerGenerator.ContainerFromIndex(i);
                if (row != null)
                {
                    rows.Add(row);
                }
            }
            return rows;
        }

        private childItem GetVisualChild<childItem>(DependencyObject obj) where childItem : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(obj); i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(obj, i);
                if (child != null && child is childItem)
                    return (childItem)child;
                else
                {
                    childItem childOfChild = GetVisualChild<childItem>(child);
                    if (childOfChild != null)
                        return childOfChild;
                }
            }
            return null;
        }
        //-----------------КОНЕЦ ПОИСК-----------------


        //-----------------ПЕЧАТЬ ДОКУМЕНТЫ-----------------

        //печать приказа
        private void order_Click(object sender, RoutedEventArgs e)
        {
            if (dataGridForm.SelectedItem == null)
            {
                MessageBox.Show("Выберите отклик!");
                return;
            }
            DataRowView selectedRow = (DataRowView)dataGridForm.SelectedItem;
            string status = selectedRow["Статус"].ToString();
            if(status == "Отклонена")
            {
                MessageBox.Show("Заявка отклонена!");
                return;
            }
            string position = selectedRow["Должность"].ToString();
            string date = selectedRow["Дата"].ToString();
            string fullName = selectedRow["ФИО Соискателя"].ToString();


            // Загрузка шаблона документа Word
            string templateFilePath = @"G:\Files\ЗаконченыеПроекты\3КУРС\4course\Дипломы\Конузелев\Документация конузелева\ПриказОприёмеНаРаботу.doc";
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Document doc = wordApp.Documents.Open(templateFilePath);
            doc.Activate();

            // Замена данных в документе
            FindAndReplace(wordApp, "[дата]", date);
            FindAndReplace(wordApp, "[дата]", date);
            FindAndReplace(wordApp, "[дата]", date);
            FindAndReplace(wordApp, "[дата]", date);
            FindAndReplace(wordApp, "[дата]", date);
            FindAndReplace(wordApp, "[дата]", date);
            FindAndReplace(wordApp, "[дата]", date);

            FindAndReplace(wordApp, "[фио]", fullName);
            FindAndReplace(wordApp, "[фио]", fullName);
            FindAndReplace(wordApp, "[должность]", position);
            FindAndReplace(wordApp, "[должность]", position);

            // Создание диалогового окна "Сохранить файл"
            Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
            saveFileDialog.Filter = "Файлы Word (*.doc)|*.doc";
            saveFileDialog.Title = "Сохранить приказ как";
            saveFileDialog.DefaultExt = "docx";

            // Сохранение документа
            if (saveFileDialog.ShowDialog() == true)
            {
                string saveFilePath = saveFileDialog.FileName;
                doc.SaveAs(saveFilePath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault);
                MessageBox.Show("Приказ успешно сохранён.");
            }

            // Закрытие документа и Word приложения
            doc.Close();
            wordApp.Quit();
        }

        //печать договора
        private void dogovor_Click(object sender, RoutedEventArgs e)
        {
            if (dataGridForm.SelectedItem == null)
            {
                MessageBox.Show("Выберите отклик!");
                return;
            }

            DataRowView selectedRow = (DataRowView)dataGridForm.SelectedItem;
            string status = selectedRow["Статус"].ToString();
            if (status == "Отклонена")
            {
                MessageBox.Show("Заявка отклонена!");
                return;
            }
            string position = selectedRow["Должность"].ToString();
            string date = selectedRow["Дата"].ToString();
            string fullName = selectedRow["ФИО Соискателя"].ToString();


            // Загрузка шаблона документа Word
            string templateFilePath = @"G:\Files\ЗаконченыеПроекты\3КУРС\4course\Дипломы\Конузелев\Документация конузелева\ТрудовойДоговор.doc";
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Document doc = wordApp.Documents.Open(templateFilePath);
            doc.Activate();

            // Замена данных в документе
            FindAndReplace(wordApp, "[дата]", date);
            FindAndReplace(wordApp, "[дата]", date);
            FindAndReplace(wordApp, "[дата]", date);
            FindAndReplace(wordApp, "[дата]", date);
            FindAndReplace(wordApp, "[дата]", date);
            FindAndReplace(wordApp, "[дата]", date);
            FindAndReplace(wordApp, "[дата]", date);

            FindAndReplace(wordApp, "[фио]", fullName);
            FindAndReplace(wordApp, "[фио]", fullName);
            FindAndReplace(wordApp, "[должность]", position);
            FindAndReplace(wordApp, "[должность]", position);

            // Создание диалогового окна "Сохранить файл"
            Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
            saveFileDialog.Filter = "Файлы Word (*.doc)|*.doc";
            saveFileDialog.Title = "Сохранить договор как";
            saveFileDialog.DefaultExt = "docx";

            // Сохранение документа
            if (saveFileDialog.ShowDialog() == true)
            {
                string saveFilePath = saveFileDialog.FileName;
                doc.SaveAs(saveFilePath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault);
                MessageBox.Show("Договор успешно сохранён.");
            }

            // Закрытие документа и Word приложения
            doc.Close();
            wordApp.Quit();
        }

        //печать личной характеристики
        private void lichHaracter_Click(object sender, RoutedEventArgs e)
        {
            if (dataGridForm.SelectedItem == null)
            {
                MessageBox.Show("Выберите соискателя!");
                return;
            }

            DataRowView selectedRow = (DataRowView)dataGridForm.SelectedItem;
            string id = selectedRow["ID"].ToString();
            string fio = selectedRow["ФИО"].ToString();
            string gender = selectedRow["Пол"].ToString();
            string date = Convert.ToDateTime(selectedRow["Дата рождения"]).ToString("dd.MM.yyyy");
            string phone = selectedRow["Телефон"].ToString();
            string lang = selectedRow["Знание языка"].ToString();

            string organization = "";
            string position = "";
            string startWorkDate = "";
            string endWorkDate = "";

            string eduInstitution = "";
            string specialty = "";
            string startEduDate = "";
            string endEduDate = "";
            string vidEducation = "";

            try
            {
                using (DatabaseConnection dbConnection = new DatabaseConnection())
                {
                    if (dbConnection.OpenConnection())
                    {
                        // Получаем данные о предыдущем опыте работы
                        string workQuery = "SELECT наименование, должность, дата_начала, дата_конца FROM ОпытРаботы WHERE id_соискателя = @id";
                        using (SqlCommand workCmd = new SqlCommand(workQuery, dbConnection.GetConnection()))
                        {
                            workCmd.Parameters.AddWithValue("@id", id);
                            using (SqlDataReader reader = workCmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    organization = reader["наименование"].ToString();
                                    position = reader["должность"].ToString();
                                    startWorkDate = Convert.ToDateTime(reader["дата_начала"]).ToString("dd.MM.yyyy");
                                    endWorkDate = Convert.ToDateTime(reader["дата_конца"]).ToString("dd.MM.yyyy");
                                }
                            }
                        }

                        // Получаем данные об образовании
                        string eduQuery = @"
    SELECT 
        o.наименование, 
        o.специальность, 
        o.номерДиплома, 
        o.дата_конца,
        vo.наименование AS ВидОбразования
    FROM Образование o
    JOIN ВидыОбразования vo ON o.id_ВидаОбразования = vo.id_ВидаОбразования
    WHERE o.id_соискателя = @id";

                        using (SqlCommand eduCmd = new SqlCommand(eduQuery, dbConnection.GetConnection()))
                        {
                            eduCmd.Parameters.AddWithValue("@id", id);
                            using (SqlDataReader reader = eduCmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    eduInstitution = reader["наименование"].ToString();
                                    specialty = reader["специальность"].ToString();
                                    startEduDate = reader["номерДиплома"].ToString();
                                    endEduDate = Convert.ToDateTime(reader["дата_конца"]).ToString("dd.MM.yyyy");
                                    vidEducation = reader["ВидОбразования"].ToString();
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Ошибка подключения к базе данных.");
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка при выполнении запроса: " + ex.Message);
                return;
            }

            // Загрузка шаблона документа Word
            //string templateFilePath = @"G:\Files\ЗаконченыеПроекты\3КУРС\4course\Дипломы\Конузелев\Документация конузелева\ЛичнаяХарактеристика.doc";
            //laptop
            string templateFilePath = @"C:\GitHub\OpenAccess\GGAEK\4course\Дипломы\Конузелев\Документация конузелева\ЛичнаяХарактеристика.doc";
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Document doc = wordApp.Documents.Open(templateFilePath);
            doc.Activate();

            // Замена данных в документе
            FindAndReplace(wordApp, "[фио]", fio);
            FindAndReplace(wordApp, "[пол]", gender);
            FindAndReplace(wordApp, "[датаРождения]", date);
            FindAndReplace(wordApp, "[телефон]", phone);
            FindAndReplace(wordApp, "[организация]", organization);
            FindAndReplace(wordApp, "[должность]", position);
            FindAndReplace(wordApp, "[датаНачалаРаботы]", startWorkDate);
            FindAndReplace(wordApp, "[датаЗавершенияРаботы]", endWorkDate);
            FindAndReplace(wordApp, "[назвУчреждения]", eduInstitution);
            FindAndReplace(wordApp, "[специальность]", specialty);
            FindAndReplace(wordApp, "[номерДиплома]", startEduDate);
            FindAndReplace(wordApp, "[датаЗавершенияУчёбы]", endEduDate);
            FindAndReplace(wordApp, "[знаниеЯзыка]", lang);
            FindAndReplace(wordApp, "[видОбразования]", vidEducation);

            // Создание диалогового окна "Сохранить файл"
            Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
            saveFileDialog.Filter = "Файлы Word (*.doc)|*.doc";
            saveFileDialog.Title = "Сохранить личную карточку сотрудника как";
            saveFileDialog.DefaultExt = "docx";

            // Сохранение документа
            if (saveFileDialog.ShowDialog() == true)
            {
                string saveFilePath = saveFileDialog.FileName;
                doc.SaveAs(saveFilePath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault);
                MessageBox.Show("Личная карточка сотрудника успешно сохранена.");
            }

            // Закрытие документа и Word приложения
            doc.Close();
            wordApp.Quit();
        }

        //метод для замены в WORD
        private void FindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object findText, object replaceText)
        {
            foreach (Range range in wordApp.ActiveDocument.StoryRanges)
            {
                range.Find.ClearFormatting();
                range.Find.Execute(FindText: findText, ReplaceWith: replaceText);
            }
        }

        private void dataGridFormSecondary_AutoGeneratedColumns(object sender, EventArgs e)
        {
            if (dataGridFormSecondary.Columns.Count > 0)
            {
                dataGridFormSecondary.Columns[0].Visibility = Visibility.Hidden;
            }
        }

        // Метод обработки изменения выбранной строки в первом DataGrid
        private void dataGridForm_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if(tableIndex != 2)
            {
                return;
            }

            if (dataGridForm.SelectedItem is DataRowView selectedRow)
            {
                // Предполагаем, что "Вид образования" - это название столбца, которое вы хотите захватить
                string vid = selectedRow["Вид образования"].ToString();
                FilterDataGridFormSecondary(vid);
            }
        }

        // Метод фильтрации данных во втором DataGrid
        private void FilterDataGridFormSecondary(string vid)
        {
            DatabaseConnection dbConnection = new DatabaseConnection();

            string querySecondary = @"
    SELECT 
        s.id_соискателя AS ID, 
        s.фио AS ФИО, 
        s.пол AS Пол, 
        FORMAT(s.дата_рождения, 'dd.MM.yyyy') AS 'Дата рождения', 
        s.телефон AS Телефон,
        s.знание_языка AS 'Знание языка' 
    FROM Соискатели s
    JOIN Образование o ON s.id_соискателя = o.id_соискателя
    JOIN ВидыОбразования vo ON o.id_ВидаОбразования = vo.id_ВидаОбразования
    WHERE vo.наименование = @vid";

            if (dbConnection.OpenConnection())
            {
                using (SqlCommand command = new SqlCommand(querySecondary, dbConnection.GetConnection()))
                {
                    command.Parameters.AddWithValue("@vid", vid);

                    try
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            DataTable dataTable = new DataTable();
                            dataTable.Load(reader);
                            dataGridFormSecondary.ItemsSource = dataTable.DefaultView;
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка при выполнении запроса: " + ex.Message);
                    }
                }
            }
            else
            {
                MessageBox.Show("Ошибка подключения к базе данных.");
            }

            // Скрываем первую колонку после загрузки данных
            if (dataGridFormSecondary.Columns.Count > 0)
            {
                dataGridFormSecondary.Columns[0].Visibility = Visibility.Hidden;
            }
        }

        private void printExselCoic_Click(object sender, RoutedEventArgs e)
        {
            if (tableIndex == 0)
            {
                MessageBox.Show("Выберите таблицу!");
                return;
            }

            string nameTable = "Подходящие соискатели";

            // Создание объекта SaveFileDialog
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Файлы Excel (*.xlsx)|*.xlsx";
            saveFileDialog.Title = "Сохранить как Excel";
            saveFileDialog.DefaultExt = "xlsx";

            if (saveFileDialog.ShowDialog() == true)
            {
                // Получение выбранного пользователем пути и имени файла
                string filePath = saveFileDialog.FileName;

                // Создание нового объекта приложения Excel
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Visible = false; // Скрываем Excel

                // Создание новой книги Excel
                Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Add(Type.Missing);

                // Создание нового листа Excel
                Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets[1];

                // Заполнение листа данными из вашего DataGrid
                for (int i = 0; i < dataGridFormSecondary.Items.Count; i++)
                {
                    var dataGridRow = (DataGridRow)dataGridFormSecondary.ItemContainerGenerator.ContainerFromIndex(i);
                    if (dataGridRow != null)
                    {
                        for (int j = 0; j < dataGridFormSecondary.Columns.Count; j++)
                        {
                            var content = dataGridFormSecondary.Columns[j].GetCellContent(dataGridRow);
                            if (content is TextBlock)
                            {
                                var text = (content as TextBlock).Text;
                                excelSheet.Cells[i + 2, j + 1] = text; // Начинаем с второй строки

                                // Если это столбец с телефонным номером, установить формат ячейки в текстовый
                                if (dataGridFormSecondary.Columns[j].Header.ToString() == "Телефон")
                                {
                                    excelSheet.Cells[i + 2, j + 1].NumberFormat = "@";
                                }
                            }
                        }
                    }
                }


                // Удаление столбца A
                Microsoft.Office.Interop.Excel.Range columnA = (Microsoft.Office.Interop.Excel.Range)excelSheet.Columns["A"];
                columnA.Delete();

                // Объединение ячеек в первой строке
                Microsoft.Office.Interop.Excel.Range headerRange = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[1, dataGridFormSecondary.Columns.Count - 1]];
                headerRange.Merge();

                // Установка текста в объединенной ячейке
                excelSheet.Cells[1, 1] = nameTable;

                // Выравнивание текста по центру и установка жирного шрифта для первой строки
                headerRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                headerRange.Font.Bold = true;

                // Добавление обводки для всей таблицы Excel
                Microsoft.Office.Interop.Excel.Range tableRange = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[dataGridFormSecondary.Items.Count + 1, dataGridFormSecondary.Columns.Count - 1]];
                tableRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                tableRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                // Выравнивание ширины столбцов
                excelSheet.UsedRange.Columns.AutoFit();

                // Сохранение книги Excel по выбранному пути
                excelBook.SaveAs(filePath);

                // Закрытие книги и приложения Excel
                excelBook.Close();
                excelApp.Quit();

                // Освобождение ресурсов COM
                Marshal.ReleaseComObject(excelSheet);
                Marshal.ReleaseComObject(excelBook);
                Marshal.ReleaseComObject(excelApp);
            }
        }


        //-----------------КОНЕЦ ПЕЧАТЬ ДОКУМЕНТЫ-----------------
    }
}
