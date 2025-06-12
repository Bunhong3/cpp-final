#ifndef FUNCTION_H
#define FUNCTION_H
#include <iostream>
#include <tabulate/table.hpp>
#include <xlnt/xlnt.hpp>
#include <exception>
#include <ctime>
using namespace std;
using namespace tabulate;
// Standard colors
#define black(text) "\033[30m" text "\033[0m"
#define red(text) "\033[31m" text "\033[0m"
#define green(text) "\033[32m" text "\033[0m"
#define yellow(text) "\033[33m" text "\033[0m"
#define blue(text) "\033[34m" text "\033[0m"
#define magenta(text) "\033[35m" text "\033[0m"
#define cyan(text) "\033[36m" text "\033[0m"
#define white(text) "\033[37m" text "\033[0m"
// Bold versions
#define bold_black(text) "\033[1;30m" text "\033[0m"
#define bold_red(text) "\033[1;31m" text "\033[0m"
#define bold_green(text) "\033[1;32m" text "\033[0m"
#define bold_yellow(text) "\033[1;33m" text "\033[0m"
#define bold_blue(text) "\033[1;34m" text "\033[0m"
#define bold_magenta(text) "\033[1;35m" text "\033[0m"
#define bold_cyan(text) "\033[1;36m" text "\033[0m"
#define bold_white(text) "\033[1;37m" text "\033[0m"

void pressEnter(){
    cout<<bold_magenta(">> Press Enter to Coutinue : ");
    cin.ignore();
}
class Client
{
private:
    string id, name, contact;

public:
    Client(string id, string name, string contact )
        : id(id), name(name), contact(contact) {}

    string getId() const { return id; }
    string getName() const { return name; }
    string getContact() const { return contact; }  
};

class Employee
{
public:
    string clientId, employeeName;
    Employee(string id, string emp) : clientId(id), employeeName(emp) {}
};

void printClientTable(const vector<Client> &clients)
{
    Table table;
    table.add_row({"Client Name", "Contact Person", "Contact Email"});
    for (auto &c : clients)
    {
        table.add_row({c.getId(), c.getName(), c.getContact()});
    }

    // Style header
    table[0].format().font_style({FontStyle::bold}).font_align(FontAlign::center).border_bottom("─");
    table[0].format().font_color(Color::cyan);
    // Style rows
    for (size_t i = 1; i < table.size(); ++i)
    {
        table[i].format().font_align(FontAlign::center);
    }

    // Global table formatting
    table.format()
        .border("│")
        .corner("+")
        .padding_top(0)
        .padding_bottom(0)
        .padding_left(1)
        .padding_right(1);

    cout << table << endl;
}

void printEmployeeTable(const vector<Employee> &employee)
{
    Table table;
    table.add_row({"Client ID", "Employee"});
    for (auto &a : employee)
    {
        table.add_row({a.clientId, a.employeeName});
    }

    table[0].format().font_style({FontStyle::bold}).font_align(FontAlign::center).border_bottom("─");
    table[0].format().font_color(Color::cyan);
    for (size_t i = 1; i < table.size(); ++i)
    {
        table[i].format().font_align(FontAlign::center);
    }
    
    cout << table << endl;
}

vector<Client> readClientsFromExcel(const string &filename)
{
    vector<Client> clients;
    xlnt::workbook wb;
    try
    {
        wb.load(filename);
        auto ws = wb.active_sheet();
        for (auto row : ws.rows(false))
        {
            if (row[0].to_string() == "Client Name")
                continue;
            Client c(
                row[0].to_string(),
                row[1].to_string(),
                row[2].to_string());
            clients.push_back(c);
        }
    }
    catch (const std::exception &e)
    {
        cout << yellow("⚠️ Error reading clients Excel file: ") << e.what() << endl;
    }
    return clients;
}

void writeClientsToExcel(const string &filename, const vector<Client> &clients)
{
    xlnt::workbook wb;
    auto ws = wb.active_sheet();
    ws.title("Clients");
    ws.cell("A1").value("Client Name");
    ws.cell("B1").value("Contact Person");
    ws.cell("C1").value("Contact Email");
    
    int row = 2;
    for (auto &c : clients)
    {
        ws.cell("A" + to_string(row)).value(c.getId());
        ws.cell("B" + to_string(row)).value(c.getName());
        ws.cell("C" + to_string(row)).value(c.getContact());
        row++;
    }
    wb.save(filename);
}

vector<Employee> reademployeeFromExcel(const string &filename)
{
    vector<Employee> employee;
    xlnt::workbook wb;
    try
    {
        wb.load(filename);
        auto ws = wb.active_sheet();
        for (auto row : ws.rows(false))
        {
            if (row[0].to_string() == "Client ID")
                continue;
            employee.emplace_back(row[0].to_string(), row[1].to_string());
        }
    }
    catch (const std::exception &e)
    {
        cout << yellow("⚠️ Error reading employee Excel file: ") << e.what() << endl;
    }
    return employee;
}

void writeemployeeToExcel(const string &filename, const vector<Employee> &employee)
{
    xlnt::workbook wb;
    auto ws = wb.active_sheet();
    ws.title("employee");
    ws.cell("A1").value("Client ID");
    ws.cell("B1").value("Employee");
    int row = 2;
    for (auto &a : employee)
    {
        ws.cell("A" + to_string(row)).value(a.clientId);
        ws.cell("B" + to_string(row)).value(a.employeeName);
        row++;
    }
    wb.save(filename);
}

// style
void printAppLogo()
{
    cout <<bold_red (R"(
   ____ _ _            _     _                                    
  / ___| (_)_ __   ___| |_  | |  
 | |   | | | '_ \ / _ \ __| | | 
 | |___| | | | | |  __/ |_  | |                       
  \____|_|_|_| |_|\___|\__| | |                                                                
    )")bold_green (R"(                        | |
                            | |  __  __                                  
                            | | |  \/  | __ _ _ __ ___   __ _  __ _  ___ 
                            | | | |\/| |/ _` | '_ ` _ \ / _` |/ _` |/ _ \
                            | | | |  | | (_| | | | | | | (_| | (_| |  __/
                            | | |_|  |_|\__,_|_| |_| |_|\__,_|\__, |\___|
                            |_|                               |___/      
    )") << "\n";    
}
void printtHeader(const std::string &title, const std::string &username = "Admin")
{
    // Get current date and time
    time_t now = time(nullptr);
    tm *ltm = localtime(&now);

    char timeBuf[32];
    strftime(timeBuf, sizeof(timeBuf), "%d-%m-%Y %H:%M:%S", ltm);

    int width = 50;
    std::string border = "+" + std::string(width, '=') + "+";

    std::cout << "\n\033[1;36m";  // Start cyan bold
    std::cout << border << "\n";

    // Centered title
    int padding = (width - title.length()) / 2;
    std::cout << "|"
              << std::string(padding, ' ') << title << std::string(width - padding - title.length(), ' ')
              << "|\n";

    std::cout << border << "\n";

    // Left-aligned user and date
    std::string userLine = "| [User]  " + username;
    userLine += std::string(width - userLine.length() + 1, ' ') + "|";
    std::cout << userLine << "\n";

    std::string dateLine = "| [Date]  " + std::string(timeBuf);
    dateLine += std::string(width - dateLine.length() + 1, ' ') + "|";
    std::cout << dateLine << "\n";

    std::cout << border << "\033[0m\n\n";  // End color
}

void printHeader(const string &title)
{
    string border = "+" + string(title.length() + 4, '=') + "+";
    cout << "\n\033[1;36m" << border << "\n";
    cout << "|  " << title << "  |\n";
    cout << border << "\033[0m\n\n";
}
#endif