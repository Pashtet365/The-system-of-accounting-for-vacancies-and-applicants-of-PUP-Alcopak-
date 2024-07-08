using Microsoft.Office.Interop.Excel;
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
using Window = System.Windows.Window;

namespace AccountingForVacanciesAndJobSeekers
{
    /// <summary>
    /// Логика взаимодействия для Вакансии.xaml
    /// </summary>
    public partial class Вакансии : Window
    {
        public Вакансии()
        {
            InitializeComponent();
            LoadActivity();
            LoadCoickateli();
            if(MainWindow.changing == 1)
            {
                mainLabel.Content = "Изменение";
                mainButton.Content = "Сохранить";
                this.Title = "Изменение данных вакансии";
                txtPost.Text = MainWindow.element2.ToString();
                txtStage.Text = MainWindow.element3.ToString();
                txtZP.Text = MainWindow.element4.ToString();
                comboActivity.Text = MainWindow.element5.ToString();
                comboCoic.Text = MainWindow.element6.ToString();
            }
        }

        private MainWindow _main;

        public Вакансии(MainWindow main) : this()
        {
            _main = main;
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            MainWindow.changing = 0;
        }

        private void mainButton_Click(object sender, RoutedEventArgs e)
        {
            // Проверка заполненности полей, кроме поля опыта работы
            if (string.IsNullOrWhiteSpace(txtPost.Text) ||
                string.IsNullOrWhiteSpace(comboActivity.Text) ||
                comboActivity.SelectedItem == null ||
                string.IsNullOrWhiteSpace(txtZP.Text) || comboCoic.SelectedItem == null)
            {
                MessageBox.Show("Заполните все поля, поле опыт работы может быть пустым!");
                return;
            }

            // Получение данных из полей
            string post = txtPost.Text;
            string activity = comboActivity.SelectedItem.ToString();
            decimal zp;
            int experience;

            // Проверка корректности ввода зарплаты
            if (!decimal.TryParse(txtZP.Text, out zp))
            {
                MessageBox.Show("Некорректный формат для заработной платы. Пожалуйста, введите число.");
                return;
            }

            // Округление зарплаты до двух знаков после запятой
            zp = Math.Round(zp, 2);

            // Получение опыта работы, если поле не пустое
            if (string.IsNullOrWhiteSpace(txtStage.Text))
            {
                experience = 0; // Если поле опыта пустое, опыт работы равен 0
            }
            else
            {
                if (!int.TryParse(txtStage.Text, out experience))
                {
                    MessageBox.Show("Некорректный формат для опыта работы. Пожалуйста, введите число.");
                    return;
                }
            }

            Coickateli selectedApplicant = (Coickateli)comboCoic.SelectedItem;

            // Ваш код для добавления или изменения записи
            try
            {
                using (DatabaseConnection dbConnection = new DatabaseConnection())
                {
                    if (dbConnection.OpenConnection())
                    {
                        if (MainWindow.changing == 0) // Добавление записи
                        {
                            // Ваш код для добавления записи
                            string insertQuery = @"
        INSERT INTO Вакансии (должность, опыт, зп, активность, дата_закрытия, id_ВидаОбразования)
        VALUES (@Position, @Experience, @ZP, @Activity, @ClosingDate, @vid)";

                            SqlCommand insertCommand = new SqlCommand(insertQuery, dbConnection.GetConnection());
                            insertCommand.Parameters.AddWithValue("@Position", post);
                            insertCommand.Parameters.AddWithValue("@Experience", experience);
                            insertCommand.Parameters.AddWithValue("@ZP", zp);
                            insertCommand.Parameters.AddWithValue("@Activity", activity);
                            insertCommand.Parameters.AddWithValue("@vid", selectedApplicant.Id);

                            if (activity == "Набор закрыт")
                            {
                                insertCommand.Parameters.AddWithValue("@ClosingDate", DateTime.Now);
                            }
                            else
                            {
                                insertCommand.Parameters.AddWithValue("@ClosingDate", DBNull.Value);
                            }

                            int rowsAffected = insertCommand.ExecuteNonQuery();

                            if (rowsAffected > 0)
                            {
                                MessageBox.Show("Запись успешно добавлена.");
                                // Очистка полей
                                txtPost.Text = "";
                                txtStage.Text = "";
                                txtZP.Text = "";
                                comboActivity.SelectedIndex = -1;
                            }
                            else
                            {
                                MessageBox.Show("Не удалось добавить запись.");
                            }
                        }
                        else // Изменение записи
                        {
                            // Ваш код для изменения записи
                            string updateQuery = @"
        UPDATE Вакансии
        SET должность = @Position, опыт = @Experience, зп = @ZP, активность = @Activity, дата_закрытия = @ClosingDate, id_ВидаОбразования = @vid
        WHERE id_вакансии = @VacancyId";

                            SqlCommand updateCommand = new SqlCommand(updateQuery, dbConnection.GetConnection());
                            updateCommand.Parameters.AddWithValue("@VacancyId", MainWindow.element1); // ID записи
                            updateCommand.Parameters.AddWithValue("@Position", post);
                            updateCommand.Parameters.AddWithValue("@Experience", experience);
                            updateCommand.Parameters.AddWithValue("@ZP", zp);
                            updateCommand.Parameters.AddWithValue("@Activity", activity);
                            updateCommand.Parameters.AddWithValue("@vid", selectedApplicant.Id);

                            if (activity == "Набор закрыт")
                            {
                                updateCommand.Parameters.AddWithValue("@ClosingDate", DateTime.Now);
                            }
                            else
                            {
                                updateCommand.Parameters.AddWithValue("@ClosingDate", DBNull.Value);
                            }

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

        public class Coickateli
        {
            public int Id { get; set; }
            public string Name { get; set; }
        }

        private void LoadCoickateli()
        {
            try
            {
                List<Coickateli> Coickateli = new List<Coickateli>();
                using (DatabaseConnection dbConnection = new DatabaseConnection())
                {
                    if (dbConnection.OpenConnection())
                    {
                        string selectQuery = "SELECT id_ВидаОбразования, наименование FROM ВидыОбразования";
                        SqlCommand selectCmd = new SqlCommand(selectQuery, dbConnection.GetConnection());
                        SqlDataReader reader = selectCmd.ExecuteReader();

                        while (reader.Read())
                        {
                            int id = reader.GetInt32(0);
                            string name = reader.GetString(1);
                            Coickateli.Add(new Coickateli { Id = id, Name = name });
                        }

                        reader.Close();
                    }
                    else
                    {
                        MessageBox.Show("Ошибка подключения к базе данных.");
                    }
                }

                // Привязать список предметов к ComboBox
                comboCoic.ItemsSource = Coickateli;
                comboCoic.DisplayMemberPath = "Name"; // Указать, какое свойство использовать для отображения в ComboBox
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }

        private void LoadActivity()
        {

            // Проверяем, что комбобокс существует и не равен null
            if (comboActivity != null)
            {
                // Очищаем комбобокс перед добавлением новых данных
                comboActivity.Items.Clear();

                // Добавляем данные в комбобокс
                comboActivity.Items.Add("Идёт набор");
                comboActivity.Items.Add("Набор закрыт");
            }
            else
            {
                MessageBox.Show("Комбобокс не найден.");
            }
        }

        //ввот только цифр
        private void txtStage_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // Проверяем, является ли введенный символ цифрой
            if (!char.IsDigit(e.Text, 0))
            {
                // Если символ не является цифрой, отменяем его ввод
                e.Handled = true;
            }
        }

        //ввод цифр с запятой
        private void txtZP_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // Проверяем, является ли введенный символ цифрой или запятой
            if (!char.IsDigit(e.Text, 0) && e.Text != ",")
            {
                // Если символ не является цифрой или запятой, отменяем его ввод
                e.Handled = true;
            }
        }
    }
}
