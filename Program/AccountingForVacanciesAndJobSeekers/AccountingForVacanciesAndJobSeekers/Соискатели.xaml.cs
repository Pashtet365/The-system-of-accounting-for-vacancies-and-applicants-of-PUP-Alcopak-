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
    /// Логика взаимодействия для Соискатели.xaml
    /// </summary>
    public partial class Соискатели : Window
    {
        public Соискатели()
        {
            InitializeComponent();
            LoadGender();
            if (MainWindow.changing == 1)
            {
                mainLabel.Content = "Изменение";
                mainButton.Content = "Сохранить";
                this.Title = "Изменение данных соискателя";
                txtName.Text = MainWindow.element2.ToString();
                comboGender.Text = MainWindow.element3.ToString();
                pickerStartDate.Text = MainWindow.element4.ToString();
                txtPhone.Text = MainWindow.element5.ToString();
                txtLang.Text = MainWindow.element6.ToString();
            }
        }

        private MainWindow _main;

        public Соискатели(MainWindow main) : this()
        {
            _main = main;
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            MainWindow.changing = 0;
        }

        private void mainButton_Click(object sender, RoutedEventArgs e)
        {
            // Проверка заполненности полей
            if (string.IsNullOrWhiteSpace(txtName.Text) ||
                string.IsNullOrWhiteSpace(comboGender.Text) ||
                pickerStartDate.SelectedDate == null ||
                string.IsNullOrWhiteSpace(txtPhone.Text) || string.IsNullOrWhiteSpace(txtLang.Text))
            {
                MessageBox.Show("Заполните все поля!");
                return;
            }

            // Проверка возраста соискателя
            DateTime currentDate = DateTime.Now;
            DateTime birthDate = pickerStartDate.SelectedDate.Value;
            int age = currentDate.Year - birthDate.Year;
            if (currentDate < birthDate.AddYears(age))
            {
                age--;
            }
            if (age < 14)
            {
                MessageBox.Show("Соискателю должно быть больше 13 лет.");
                return;
            }

            if (age > 99)
            {
                MessageBox.Show("Нельзя добавить соискателя возрастом 99 лет.");
                return;
            }

            // Получение данных из полей
            string name = txtName.Text;
            string gender = comboGender.Text;
            DateTime dateOfBirth = pickerStartDate.SelectedDate.Value;
            string phone = txtPhone.Text;

            try
            {
                using (DatabaseConnection dbConnection = new DatabaseConnection())
                {
                    if (dbConnection.OpenConnection())
                    {
                        if (MainWindow.changing == 0) // Добавление записи
                        {
                            string insertQuery = @"
                        INSERT INTO Соискатели (фио, пол, дата_рождения, телефон, знание_языка)
                        VALUES (@Name, @Gender, @DateOfBirth, @Phone, @lang)";

                            SqlCommand insertCommand = new SqlCommand(insertQuery, dbConnection.GetConnection());
                            insertCommand.Parameters.AddWithValue("@Name", name);
                            insertCommand.Parameters.AddWithValue("@Gender", gender);
                            insertCommand.Parameters.AddWithValue("@DateOfBirth", dateOfBirth);
                            insertCommand.Parameters.AddWithValue("@Phone", phone);
                            insertCommand.Parameters.AddWithValue("@lang", txtLang.Text);

                            int rowsAffected = insertCommand.ExecuteNonQuery();

                            if (rowsAffected > 0)
                            {
                                MessageBox.Show("Запись успешно добавлена.");
                                // Очистка полей
                                txtName.Text = "";
                                comboGender.Text = "";
                                pickerStartDate.SelectedDate = null;
                                txtPhone.Text = "";
                                txtLang.Text = "";
                            }
                            else
                            {
                                MessageBox.Show("Не удалось добавить запись.");
                            }
                        }
                        else // Изменение записи
                        {
                            // Реализация изменения записи
                            string updateQuery = @"
                        UPDATE Соискатели
                        SET фио = @Name, пол = @Gender, дата_рождения = @DateOfBirth, телефон = @Phone, знание_языка = @lang
                        WHERE id_соискателя = @Id";

                            SqlCommand updateCommand = new SqlCommand(updateQuery, dbConnection.GetConnection());
                            updateCommand.Parameters.AddWithValue("@Id", MainWindow.element1); // ID записи
                            updateCommand.Parameters.AddWithValue("@Name", name);
                            updateCommand.Parameters.AddWithValue("@Gender", gender);
                            updateCommand.Parameters.AddWithValue("@DateOfBirth", dateOfBirth);
                            updateCommand.Parameters.AddWithValue("@Phone", phone);
                            updateCommand.Parameters.AddWithValue("@lang", txtLang.Text);

                            int rowsAffected = updateCommand.ExecuteNonQuery();

                            if (rowsAffected > 0)
                            {
                                MessageBox.Show("Запись успешно изменена.");
                            }
                            else
                            {
                                MessageBox.Show("Не удалось изменить запись.");
                            }
                        }

                        // Обновление данных в таблице (если необходимо)
                        _main.Refresh(sender, e);
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

        private void LoadGender()
        {

            // Проверяем, что комбобокс существует и не равен null
            if (comboGender != null)
            {
                // Очищаем комбобокс перед добавлением новых данных
                comboGender.Items.Clear();

                // Добавляем данные в комбобокс
                comboGender.Items.Add("Мужской");
                comboGender.Items.Add("Женский");
            }
            else
            {
                MessageBox.Show("Комбобокс не найден.");
            }
        }

        //ввод только букв
        private void txtName_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // Получаем введенный символ
            char inputChar = e.Text[0];

            // Проверяем, является ли символ буквой
            if (!char.IsLetter(inputChar))
            {
                // Если символ не является буквой, отменяем его ввод
                e.Handled = true;
            }
        }

        //ввод телефона
        private void txtPhone_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // Получаем текущий текст в текстовом поле
            string currentText = txtPhone.Text;

            // Получаем вводимый символ
            char inputChar = e.Text[0];

            // Проверяем, является ли вводимый символ цифрой или плюсом
            if (!char.IsDigit(inputChar) && inputChar != '+')
            {
                // Отменяем ввод, если символ не является цифрой или плюсом
                e.Handled = true;
            }

            // Проверяем, чтобы символ "+" вводился только в начале строки
            if (inputChar == '+' && currentText.Length > 0)
            {
                e.Handled = true;
            }

            // Проверяем, чтобы первые три символа были "+375"
            if (currentText.Length == 0 && inputChar != '+')
            {
                e.Handled = true;
            }
            else if (currentText.Length == 1 && inputChar != '3')
            {
                e.Handled = true;
            }
            else if (currentText.Length == 2 && inputChar != '7')
            {
                e.Handled = true;
            }

            // Проверяем, чтобы количество цифр не превышало 13 (включая "+")
            if (currentText.Length >= 13)
            {
                e.Handled = true;
            }
        }
    }
}
