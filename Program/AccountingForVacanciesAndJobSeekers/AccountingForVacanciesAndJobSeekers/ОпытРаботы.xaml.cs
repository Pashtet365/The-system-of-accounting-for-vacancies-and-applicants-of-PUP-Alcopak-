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

namespace AccountingForVacanciesAndJobSeekers
{
    /// <summary>
    /// Логика взаимодействия для ОпытРаботы.xaml
    /// </summary>
    public partial class ОпытРаботы : Window
    {
        public ОпытРаботы()
        {
            InitializeComponent();
            LoadCoickateli();
            if (MainWindow.changing == 1)
            {
                mainLabel.Content = "Изменение";
                mainButton.Content = "Сохранить";
                this.Title = "Изменение данных опыта работы";
                comboCoic.Text = MainWindow.element2.ToString();
                txtName.Text = MainWindow.element3.ToString();
                txtPost.Text = MainWindow.element4.ToString();
                pickerStartDate.Text = MainWindow.element5.ToString();
                pickerLastDate.Text = MainWindow.element6.ToString();
            }
        }

        private MainWindow _main;

        public ОпытРаботы(MainWindow main) : this()
        {
            _main = main;
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            MainWindow.changing = 0;
        }

        private void mainButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(comboCoic.Text) || string.IsNullOrEmpty(txtName.Text) || string.IsNullOrEmpty(txtPost.Text) || string.IsNullOrEmpty(pickerStartDate.Text) || string.IsNullOrEmpty(pickerLastDate.Text))
            {
                MessageBox.Show("Заполните все поля!");
                return;
            }

            DateTime startDate, lastDate;
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

            try
            {
                using (DatabaseConnection dbConnection = new DatabaseConnection())
                {
                    if (dbConnection.OpenConnection())
                    {
                        if (MainWindow.changing == 0) // Добавление записи
                        {
                            string insertQuery = @"
                        INSERT INTO ОпытРаботы (id_соискателя, наименование, должность, дата_начала, дата_конца) 
                        VALUES (@idCoic, @name, @post, @startDate, @lastDate)";

                            SqlCommand insertCommand = new SqlCommand(insertQuery, dbConnection.GetConnection());

                            // Добавляем параметры к запросу
                            insertCommand.Parameters.AddWithValue("@idCoic", ((Coickateli)comboCoic.SelectedItem).Id);
                            insertCommand.Parameters.AddWithValue("@name", txtName.Text);
                            insertCommand.Parameters.AddWithValue("@post", txtPost.Text);
                            insertCommand.Parameters.AddWithValue("@startDate", startDate);
                            insertCommand.Parameters.AddWithValue("@lastDate", lastDate);

                            // Выполняем запрос
                            int rowsAffected = insertCommand.ExecuteNonQuery();

                            if (rowsAffected > 0)
                            {
                                MessageBox.Show("Запись успешно добавлена.");
                                comboCoic.SelectedIndex = -1;
                                txtName.Text = "";
                                txtPost.Text = "";
                                pickerLastDate.Text = "";
                                pickerStartDate.Text = "";
                            }
                            else
                            {
                                MessageBox.Show("Ошибка при добавлении записи.");
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

                            string updateQuery = @"
    UPDATE ОпытРаботы 
    SET id_соискателя = @idCoic, наименование = @name, должность = @post, дата_начала = @startDate, дата_конца = @lastDate 
    WHERE id_опыта_работы = @orderId";

                            SqlCommand updateCommand = new SqlCommand(updateQuery, dbConnection.GetConnection());

                            // Добавляем параметры к запросу
                            updateCommand.Parameters.AddWithValue("@idCoic", ((Coickateli)comboCoic.SelectedItem).Id);
                            updateCommand.Parameters.AddWithValue("@name", txtName.Text);
                            updateCommand.Parameters.AddWithValue("@post", txtPost.Text);
                            updateCommand.Parameters.AddWithValue("@startDate", startDate);
                            updateCommand.Parameters.AddWithValue("@lastDate", lastDate);
                            updateCommand.Parameters.AddWithValue("@orderId", MainWindow.element1);

                            // Выполняем запрос
                            int rowsAffected = updateCommand.ExecuteNonQuery();

                            if (rowsAffected > 0)
                            {
                                MessageBox.Show("Запись успешно изменена.");
                                // Очистка полей формы или обновление данных в интерфейсе
                            }
                            else
                            {
                                MessageBox.Show("Не удалось изменить запись.");
                            }
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
                        FROM ОпытРаботы o
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

        private void txtName_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // Проверяем, является ли введенный символ буквой
            if (!char.IsLetter(e.Text, 0))
            {
                e.Handled = true; // Отменяем ввод, если символ не является буквой
            }
        }

        private void txtPost_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // Проверяем, является ли введенный символ буквой
            if (!char.IsLetter(e.Text, 0))
            {
                e.Handled = true; // Отменяем ввод, если символ не является буквой
            }
        }
    }
}
