using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
    /// Логика взаимодействия для Образование.xaml
    /// </summary>
    public partial class Образование : Window
    {
        public Образование()
        {
            InitializeComponent();
            LoadCoickateli();
            LoadCoickateli1();
            if (MainWindow.changing == 1)
            {
                mainLabel.Content = "Изменение";
                mainButton.Content = "Сохранить";
                this.Title = "Изменение данных образования";
                comboCoic.Text = MainWindow.element2.ToString();
                txtName.Text = MainWindow.element3.ToString();
                txtPost.Text = MainWindow.element4.ToString();
                txtDiplom.Text = MainWindow.element5.ToString();
                pickerLastDate.Text = MainWindow.element6.ToString();
                comboCoic1.Text = MainWindow.element7.ToString();
            }
        }

        private MainWindow _main;

        public Образование(MainWindow main) : this()
        {
            _main = main;
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            MainWindow.changing = 0;
        }

        private void mainButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(comboCoic.Text) || string.IsNullOrEmpty(txtName.Text) || string.IsNullOrEmpty(txtPost.Text) || string.IsNullOrEmpty(txtDiplom.Text) || string.IsNullOrEmpty(pickerLastDate.Text))
            {
                MessageBox.Show("Заполните все поля!");
                return;
            }

            DateTime? lastDate = pickerLastDate.SelectedDate;

            try
            {
                using (DatabaseConnection dbConnection = new DatabaseConnection())
                {
                    if (dbConnection.OpenConnection())
                    {
                        if (MainWindow.changing == 0) // Добавление записи
                        {
                            string insertQuery = @"
                        INSERT INTO Образование (id_соискателя, наименование, специальность, номерДиплома, дата_конца, id_ВидаОбразования) 
                        VALUES (@idCoic, @name, @post, @startDate, @lastDate, @vid)";

                            SqlCommand insertCommand = new SqlCommand(insertQuery, dbConnection.GetConnection());

                            // Добавляем параметры к запросу
                            insertCommand.Parameters.AddWithValue("@idCoic", ((Coickateli)comboCoic.SelectedItem).Id);
                            insertCommand.Parameters.AddWithValue("@name", txtName.Text);
                            insertCommand.Parameters.AddWithValue("@post", txtPost.Text);
                            insertCommand.Parameters.AddWithValue("@startDate", txtDiplom.Text);
                            insertCommand.Parameters.AddWithValue("@lastDate", lastDate);
                            insertCommand.Parameters.AddWithValue("@vid", ((Coickateli1)comboCoic1.SelectedItem).Id);

                            // Выполняем запрос
                            int rowsAffected = insertCommand.ExecuteNonQuery();

                            if (rowsAffected > 0)
                            {
                                MessageBox.Show("Запись успешно добавлена.");
                                comboCoic.SelectedIndex = -1;
                                txtName.Text = "";
                                txtPost.Text = "";
                                pickerLastDate.Text = "";
                                txtDiplom.Text = "";
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
    UPDATE Образование 
    SET id_соискателя = @idCoic, наименование = @name, специальность = @post, номерДиплома = @startDate, дата_конца = @lastDate, id_ВидаОбразования = @vid 
    WHERE id_образования = @orderId";

                            SqlCommand updateCommand = new SqlCommand(updateQuery, dbConnection.GetConnection());

                            // Добавляем параметры к запросу
                            updateCommand.Parameters.AddWithValue("@idCoic", ((Coickateli)comboCoic.SelectedItem).Id);
                            updateCommand.Parameters.AddWithValue("@name", txtName.Text);
                            updateCommand.Parameters.AddWithValue("@post", txtPost.Text);
                            updateCommand.Parameters.AddWithValue("@startDate", txtDiplom.Text);
                            updateCommand.Parameters.AddWithValue("@lastDate", lastDate);
                            updateCommand.Parameters.AddWithValue("@vid", ((Coickateli1)comboCoic1.SelectedItem).Id);
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
                        FROM Образование o
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

        public class Coickateli1
        {
            public int Id { get; set; }
            public string Name { get; set; }
        }

        private void LoadCoickateli1()
        {
            try
            {
                List<Coickateli1> CoickateliList1 = new List<Coickateli1>();
                using (DatabaseConnection dbConnection = new DatabaseConnection())
                {
                    if (dbConnection.OpenConnection())
                    {
                        string selectQuery = @"
                    SELECT id_ВидаОбразования, наименование 
                    FROM ВидыОбразования";

                        SqlCommand selectCmd = new SqlCommand(selectQuery, dbConnection.GetConnection());
                        SqlDataReader reader = selectCmd.ExecuteReader();

                        while (reader.Read())
                        {
                            int id = reader.GetInt32(0);
                            string name = reader.GetString(1);
                            CoickateliList1.Add(new Coickateli1 { Id = id, Name = name });
                        }

                        reader.Close();
                    }
                    else
                    {
                        MessageBox.Show("Ошибка подключения к базе данных.");
                    }
                }

                // Привязать список соискателей к ComboBox
                comboCoic1.ItemsSource = CoickateliList1;
                comboCoic1.DisplayMemberPath = "Name"; // Указать, какое свойство использовать для отображения в ComboBox
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

        private void txtDiplom_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // Используем регулярное выражение для проверки, что вводимый символ - цифра
            Regex regex = new Regex("[^0-9]+"); // Нецифровые символы
            e.Handled = regex.IsMatch(e.Text); // Если символ не соответствует шаблону, отменяем ввод
        }
    }
}
