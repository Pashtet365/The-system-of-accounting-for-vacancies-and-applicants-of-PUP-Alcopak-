using System;
using System.Collections.Generic;
using System.Data.SqlClient;
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
using System.Xml.Linq;

namespace AccountingForVacanciesAndJobSeekers
{
    /// <summary>
    /// Логика взаимодействия для ВоинскийУчёт.xaml
    /// </summary>
    public partial class ВоинскийУчёт : Window
    {
        public ВоинскийУчёт()
        {
            InitializeComponent();
            LoadCoickateli();
            LoadStatus();
            if (MainWindow.changing == 1)
            {
                mainLabel.Content = "Изменение";
                mainButton.Content = "Сохранить";
                this.Title = "Изменение данных опыта работы";
                comboCoic.Text = MainWindow.element2.ToString();
                comboGodnost.Text = MainWindow.element3.ToString();
                pickerStartDate.Text = MainWindow.element4.ToString();
                pickerLastDate.Text = MainWindow.element5.ToString();
            }
        }

        private MainWindow _main;

        public ВоинскийУчёт(MainWindow main) : this()
        {
            _main = main;
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            MainWindow.changing = 0;
        }

        private void mainButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(comboCoic.Text) || string.IsNullOrEmpty(comboGodnost.Text))
            {
                MessageBox.Show("Выберите соискателя и годность!");
                return;
            }

            DateTime startDate = DateTime.MinValue;
            DateTime lastDate = DateTime.MinValue;

            if (comboGodnost.Text != "Годен" && comboGodnost.Text != "Годен с ограничениями" && comboGodnost.Text != "Не годен и снимается с воинского учёта")
            {
                if (string.IsNullOrWhiteSpace(pickerStartDate.Text) || string.IsNullOrWhiteSpace(pickerLastDate.Text))
                {
                    MessageBox.Show("Поля дат не могут быть пустыми!");
                    return;
                }

                if (!DateTime.TryParse(pickerStartDate.Text, out startDate) || !DateTime.TryParse(pickerLastDate.Text, out lastDate))
                {
                    MessageBox.Show("Неверный формат даты!");
                    return;
                }

                if (startDate > lastDate)
                {
                    MessageBox.Show("Дата начала не может быть позже даты окончания!");
                    return;
                }
            }

            try
            {
                using (DatabaseConnection dbConnection = new DatabaseConnection())
                {
                    if (dbConnection.OpenConnection())
                    {
                        string query = string.Empty;

                        if (MainWindow.changing == 0) // Добавление записи
                        {
                            if (startDate == DateTime.MinValue || lastDate == DateTime.MinValue)
                            {
                                // Вставка данных без даты
                                query = @"
                            INSERT INTO ВоинскийУчёт (id_соискателя, годность) 
                            VALUES (@idCoic, @post)";
                            }
                            else
                            {
                                // Вставка данных с датой
                                query = @"
                            INSERT INTO ВоинскийУчёт (id_соискателя, годность, дата_начала, дата_конца) 
                            VALUES (@idCoic, @post, @startDate, @lastDate)";
                            }
                        }
                        else // Изменение записи
                        {
                            // Проверяем, что выбран элемент в ComboBox
                            if (comboCoic.SelectedItem == null || string.IsNullOrEmpty(((Coickateli)comboCoic.SelectedItem).Id.ToString()))
                            {
                                MessageBox.Show("Выберите соискателя!");
                                return;
                            }

                            if (startDate == DateTime.MinValue || lastDate == DateTime.MinValue)
                            {
                                // Изменение данных без даты
                                query = @"
                            UPDATE ВоинскийУчёт 
                            SET id_соискателя = @idCoic, годность = @post 
                            WHERE id_военки = @orderId";
                            }
                            else
                            {
                                // Изменение данных с датой
                                query = @"
                            UPDATE ВоинскийУчёт 
                            SET id_соискателя = @idCoic, годность = @post, дата_начала = @startDate, дата_конца = @lastDate 
                            WHERE id_военки = @orderId";
                            }
                        }

                        SqlCommand command = new SqlCommand(query, dbConnection.GetConnection());

                        // Добавляем параметры к запросу
                        command.Parameters.AddWithValue("@idCoic", ((Coickateli)comboCoic.SelectedItem).Id);
                        command.Parameters.AddWithValue("@post", comboGodnost.Text);
                        if (startDate != DateTime.MinValue && lastDate != DateTime.MinValue)
                        {
                            command.Parameters.AddWithValue("@startDate", startDate);
                            command.Parameters.AddWithValue("@lastDate", lastDate);
                        }
                        if (MainWindow.changing == 1)
                        {
                            command.Parameters.AddWithValue("@orderId", MainWindow.element1);
                        }

                        // Выполняем запрос
                        int rowsAffected = command.ExecuteNonQuery();

                        if (rowsAffected > 0)
                        {
                            if (MainWindow.changing == 0)
                            {
                                MessageBox.Show("Запись успешно добавлена.");
                            }
                            else
                            {
                                MessageBox.Show("Запись успешно изменена.");
                            }

                            // Очистка полей формы или обновление данных в интерфейсе
                            comboCoic.SelectedIndex = -1;
                            comboGodnost.SelectedIndex = -1;
                            pickerLastDate.Text = "";
                            pickerStartDate.Text = "";
                        }
                        else
                        {
                            MessageBox.Show("Ошибка при " + (MainWindow.changing == 0 ? "добавлении" : "изменении") + " записи.");
                        }

                        // Обновление данных в таблице (если необходимо)
                        _main.Refresh(sender, e);
                        comboCoic.ItemsSource = null; // Очищаем источник данных
                        LoadCoickateli(); // Перезагружаем данные
                    }
                    else
                    {
                        MessageBox.Show("Ошибка подключения к базе данных.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка при выполнении запроса: " + ex.Message);
            }
        }

        public class Coickateli
        {
            public int Id { get; set; }
            public string Name { get; set; }
        }

        private void LoadCoickateli()
        {
            try
            {
                List<Coickateli> CoickateliList = new List<Coickateli>();
                using (DatabaseConnection dbConnection = new DatabaseConnection())
                {
                    if (dbConnection.OpenConnection())
                    {
                        string selectQuery = @"
                    SELECT id_соискателя, фио 
                    FROM Соискатели s
                    WHERE NOT EXISTS (
                        SELECT 1
                        FROM ВоинскийУчёт o
                        WHERE s.id_соискателя = o.id_соискателя
                    )";

                        SqlCommand selectCmd = new SqlCommand(selectQuery, dbConnection.GetConnection());
                        SqlDataReader reader = selectCmd.ExecuteReader();

                        while (reader.Read())
                        {
                            int id = reader.GetInt32(0);
                            string name = reader.GetString(1);
                            CoickateliList.Add(new Coickateli { Id = id, Name = name });
                        }

                        reader.Close();
                    }
                    else
                    {
                        MessageBox.Show("Ошибка подключения к базе данных.");
                    }
                }

                // Привязать список соискателей к ComboBox
                comboCoic.ItemsSource = CoickateliList;
                comboCoic.DisplayMemberPath = "Name"; // Указать, какое свойство использовать для отображения в ComboBox
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }

        private void LoadStatus()
        {

            // Проверяем, что комбобокс существует и не равен null
            if (comboGodnost != null)
            {
                // Очищаем комбобокс перед добавлением новых данных
                comboGodnost.Items.Clear();

                // Добавляем данные в комбобокс
                comboGodnost.Items.Add("Годен");
                comboGodnost.Items.Add("Годен с ограничениями");
                comboGodnost.Items.Add("Временно не годен");
                comboGodnost.Items.Add("Годен к прохождению военной службы вне строя в мирное время");
                comboGodnost.Items.Add("Не годен в мирное время");
                comboGodnost.Items.Add("Не годен и снимается с воинского учёта");
            }
            else
            {
                MessageBox.Show("Комбобокс не найден.");
            }
        }
    }
}
