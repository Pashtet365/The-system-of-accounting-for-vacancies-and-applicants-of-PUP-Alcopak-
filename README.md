# The-system-of-accounting-for-vacancies-and-applicants-of-PUP-Alcopak-

# Troubleshooting errors
____
:black_square_button:To work with the program, you need to change the path to the database, since it is fixed. To do this, edit the "SqlConnection.cs" file. It is necessary to specify a valid path to the database. 
```
connectionString = "Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=\"G:\\Files\\Program\\DB.mdf\";Integrated Security=True";
```
____
:black_square_button:Fixed paths are also set for printing documents.
To change them, you need to edit the file "MainWindow.xaml.cs". 
Where in the document packaging methods it is necessary to change the variables 
```
"string template File Path = @"\SourceDocuments\file.type ";"
```
____
:black_square_button:If the program requires Material Design, then it must be installed via NuGet packages.

Material Design Colors 2.1.4
Material Design Themes 4.9.0

And also Microsoft.Xmal.Behaviors.Wpf 1.1.39