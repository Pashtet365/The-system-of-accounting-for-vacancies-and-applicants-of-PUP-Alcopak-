using System;
using System.Data;
using System.Data.SqlClient;

public class DatabaseConnection : IDisposable
{
    private string connectionString;
    private SqlConnection connection;

    public DatabaseConnection()
    {
        //computer
        connectionString = "Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=\"G:\\Files\\ЗаконченыеПроекты\\3КУРС\\4course\\Дипломы\\Конузелев\\Программа Конузелева\\AccountingForVacanciesAndJobSeekers\\AccountingForVacanciesAndJobSeekers\\DB.mdf\"; Integrated Security=True";
        //laptor
        //connectionString = "Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=\"C:\\GitHub\\OpenAccess\\GGAEK\\4course\\Дипломы\\Конузелев\\Программа Конузелева\\AccountingForVacanciesAndJobSeekers\\AccountingForVacanciesAndJobSeekers\\DB.mdf\";Integrated Security=True";
        connection = new SqlConnection(connectionString);
    }

    public bool OpenConnection()
    {
        try
        {
            if (connection.State == ConnectionState.Closed)
            {
                connection.Open();
            }
            return true;
        }
        catch (SqlException)
        {
            // Обработка ошибки подключения
            return false;
        }
    }

    public bool CloseConnection()
    {
        try
        {
            if (connection.State == ConnectionState.Open)
            {
                connection.Close();
            }
            return true;
        }
        catch (SqlException)
        {
            // Обработка ошибки закрытия подключения
            return false;
        }
    }

    public SqlConnection GetConnection()
    {
        return connection;
    }

    public void Dispose()
    {
        // Закрыть соединение при уничтожении объекта
        if (connection.State == ConnectionState.Open)
        {
            connection.Close();
        }
    }
}