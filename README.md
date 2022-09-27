# Madden Player Stats Importer

This program gathers season/career stats from Pro Football Reference/Stathead for use in Madden.
### Releases
The current release for this project can be found [here](https://github.com/NoahJele/madden-stats-profootballreference/releases/). It includes the readme file and the exe needed to run the program.

### The Basics

To run this program, take the maddenimporter.excel.exe file. If you double click it, it'll run the current season's stats for active NFL players and will output to an Excel file.
To run with parameters or to run the program with the goal of gathering career stats, open a Command Prompt, cd to the folder where the maddenimporter.excel.exe file is, and run maddenimporter.excel.exe -h to see available options.

### Login

For career stats, this program pulls from [StatHead](https://stathead.com), and it requires a paid subscription to access the data. You can provide your username and password in two ways:

1. Use the `-u (username)` and `-p (password)` command-line arguments while running the program; or
2. Create a file in the root directory called `login.private` with two lines. The first line is your username, and the second line is your password.

**The program will not save your username and / or password. You must either create a file or enter the information every time you run the program.**

### From source

If you want to save the data as an Excel file, use

```bash
dotnet run -p MaddenImporter.Excel/ -- [-y year] [--path path] [--career] [-u username] [-p password]
```

If you wish to use a different file format, the `MaddenImporter` class library contains all the necessary functions and data.

### From dist

The executable file for your operating system can be found in the `dist/` folder, and in the main folder. Currently only Windows is supported.

```bash
./maddenimporter.excel(.exe) [-y year] [--path path] [--career] [-u username] [-p password]
```
