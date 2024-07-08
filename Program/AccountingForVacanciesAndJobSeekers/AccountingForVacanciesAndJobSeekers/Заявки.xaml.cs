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
    /// Логика взаимодействия для Заявки.xaml
    /// </summary>
    public partial class Заявки : Window
    {
        public Заявки()
        {
            InitializeComponent();
            LoadStatus();
            LoadVacancii();
            LoadCoickateli();
            if (MainWindow.changing == 1)
            {
                mainLabel.Content = "Изменение";
                mainButton.Content = "Сохранить";
                this.Title = "Изменение данных отклика";
                comboVac.Text = MainWindow.element2.ToString();
                comboCoic.Text = MainWindow.element3.ToString();
                comboStatus.Text = MainWindow.element4.ToString();
            }
        }

        private MainWindow _main;

        public Заявки(MainWindow main) : this()
        {
            _main = main;
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            MainWindow.changing = 0;
        }

        private void mainButton_Click(object sender, RoutedEventArgs e)
        {
            // Проверка заполненности всех полей
            if (comboVac.SelectedItem == null || comboCoic.SelectedItem == null || comboStatus.SelectedItem == null)
            {
                MessageBox.Show("Заполните все поля!");
                return;
            }

            // Получение выбранных значений из комбо-боксов
            Vacancii selectedVacancy = (Vacancii)comboVac.SelectedItem;
            Coickateli selectedApplicant = (Coickateli)comboCoic.SelectedItem;
            string selectedStatus = comboStatus.SelectedItem.ToString();

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
        INSERT INTO Заявки (id_вакансии, id_соискателя, статус, дата)
        VALUES (@VacancyId, @ApplicantId, @Status, @Date)";

                            SqlCommand insertCommand = new SqlCommand(insertQuery, dbConnection.GetConnection());
                            insertCommand.Parameters.AddWithValue("@VacancyId", selectedVacancy.Id);
                            insertCommand.Parameters.AddWithValue("@ApplicantId", selectedApplicant.Id);
                            insertCommand.Parameters.AddWithValue("@Status", selectedStatus);
                            insertCommand.Parameters.AddWithValue("@Date", DateTime.Now);

                            int rowsAffected = insertCommand.ExecuteNonQuery();

                            if (rowsAffected > 0)
                            {
                                MessageBox.Show("Запись успешно добавлена.");
                                // Очистка полей
                                comboVac.SelectedIndex = -1;
                                comboCoic.SelectedIndex = -1;
                                comboStatus.SelectedIndex = -1;
                            }
                            else
                            {
                                MessageBox.Show("Не удалось добавить запись.");
                            }
                        }
                        else // Изменение записи
                        {
                            // Проверка выбранной записи для изменения
                            if (selectedVacancy == null || selectedApplicant == null || string.IsNullOrWhiteSpace(selectedStatus))
                            {
                                MessageBox.Show("Выберите запись для изменения!");
                                return;
                            }

                            try
                            {
                                // Ваш код для изменения записи
                                string updateQuery = @"
            UPDATE Заявки
            SET id_вакансии = @VacancyId, id_соискателя = @ApplicantId, статус = @Status, дата = @Date
            WHERE id_заявки = @ApplicationId";

                                SqlCommand updateCommand = new SqlCommand(updateQuery, dbConnection.GetConnection());
                                updateCommand.Parameters.AddWithValue("@VacancyId", selectedVacancy.Id);
                                updateCommand.Parameters.AddWithValue("@ApplicantId", selectedApplicant.Id);
                                updateCommand.Parameters.AddWithValue("@Status", selectedStatus);
                                updateCommand.Parameters.AddWithValue("@ApplicationId", MainWindow.element1);
                                updateCommand.Parameters.AddWithValue("@Date", DateTime.Now);

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
                            catch (Exception ex)
                            {
                                MessageBox.Show("Произошла ошибка при выполнении запроса: " + ex.Message);
                            }
                        }
                        // Проверка статуса заявок и обновление вакансии
                        UpdateVacancyStatusIfAccepted(dbConnection, selectedVacancy.Id, selectedStatus);

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

        private void UpdateVacancyStatusIfAccepted(DatabaseConnection dbConnection, int vacancyId, string selectedStatus)
        {
            // Проверяем статус заявки
            if (selectedStatus == "Принята")
            {
                // Обновляем статус вакансии
                string updateVacancyQuery = @"
            UPDATE Вакансии
            SET активность = N'Набор закрыт', дата_закрытия = @ClosingDate
            WHERE id_вакансии = @VacancyId";

                SqlCommand updateVacancyCommand = new SqlCommand(updateVacancyQuery, dbConnection.GetConnection());
                updateVacancyCommand.Parameters.AddWithValue("@VacancyId", vacancyId);
                updateVacancyCommand.Parameters.AddWithValue("@ClosingDate", DateTime.Now);

                updateVacancyCommand.ExecuteNonQuery();
            }
        }

        public class Vacancii
        {
            public int Id { get; set; }
            public string Name { get; set; }
        }

        private void LoadVacancii()
        {
            try
            {
                List<Vacancii> Vacancii = new List<Vacancii>();
                using (DatabaseConnection dbConnection = new DatabaseConnection())
                {
                    if (dbConnection.OpenConnection())
                    {
                        string selectQuery = @"
                    SELECT id_вакансии, должность 
                    FROM Вакансии 
                    WHERE активность != N'Набор закрыт'";
                        SqlCommand selectCmd = new SqlCommand(selectQuery, dbConnection.GetConnection());
                        SqlDataReader reader = selectCmd.ExecuteReader();

                        while (reader.Read())
                        {
                            int id = reader.GetInt32(0);
                            string name = reader.GetString(1);
                            Vacancii.Add(new Vacancii { Id = id, Name = name });
                        }

                        reader.Close();
                    }
                    else
                    {
                        MessageBox.Show("Ошибка подключения к базе данных.");
                    }
                }

                // Привязать список предметов к ComboBox
                comboVac.ItemsSource = Vacancii;
                comboVac.DisplayMemberPath = "Name"; // Указать, какое свойство использовать для отображения в ComboBox
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
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
                        string selectQuery = "SELECT id_соискателя, фио FROM Соискатели";
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

        private void LoadStatus()
        {

            // Проверяем, что комбобокс существует и не равен null
            if (comboStatus != null)
            {
                // Очищаем комбобокс перед добавлением новых данных
                comboStatus.Items.Clear();

                // Добавляем данные в комбобокс
                comboStatus.Items.Add("Принята");
                comboStatus.Items.Add("Отклонена");
            }
            else
            {
                MessageBox.Show("Комбобокс не найден.");
            }
        }
    }
}
