#ifndef FUNCTION_H
#define FUNCTION_H
#include <iostream>
#include <tabulate/table.hpp>
#include <xlnt/xlnt.hpp>
#include <exception>
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
    cout<<"Press Enter to Coutinue : ";
    cin.ignore();
}
class Client
{
private:
    string id, name, contact, company, address, service;

public:
    Client(string id, string name, string contact, string company, string address, string service)
        : id(id), name(name), contact(contact), company(company), address(address), service(service) {}

    string getId() const { return id; }
    string getName() const { return name; }
    string getContact() const { return contact; }
    string getCompany() const { return company; }
    string getAddress() const { return address; }
    string getService() const { return service; }
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
    table.add_row({"ID", "Name", "Contact", "Company", "Address", "Service"});
    for (auto &c : clients)
    {
        table.add_row({c.getId(), c.getName(), c.getContact(), c.getCompany(), c.getAddress(), c.getService()});
    }

    // Style header
    table[0].format().font_style({FontStyle::bold}).font_align(FontAlign::center).border_bottom("─");

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

    for (size_t i = 1; i < table.size(); ++i)
    {
        table[i].format().font_align(FontAlign::center);
    }
    table.format()
        .border("│")
        .corner("+")
        .padding_left(1)
        .padding_right(1)
        .border_top("═")
        .border_bottom("═")
        .border_left("║")
        .border_right("║");

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
            if (row[0].to_string() == "ID")
                continue;
            Client c(
                row[0].to_string(),
                row[1].to_string(),
                row[2].to_string(),
                row[3].to_string(),
                row[4].to_string(),
                row[5].to_string());
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
    ws.cell("A1").value("ID");
    ws.cell("B1").value("Name");
    ws.cell("C1").value("Contact");
    ws.cell("D1").value("Company");
    ws.cell("E1").value("Address");
    ws.cell("F1").value("Service");
    int row = 2;
    for (auto &c : clients)
    {
        ws.cell("A" + to_string(row)).value(c.getId());
        ws.cell("B" + to_string(row)).value(c.getName());
        ws.cell("C" + to_string(row)).value(c.getContact());
        ws.cell("D" + to_string(row)).value(c.getCompany());
        ws.cell("E" + to_string(row)).value(c.getAddress());
        ws.cell("F" + to_string(row)).value(c.getService());
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

void deleteClient(vector<Client> &clients, vector<Employee> &employee,
                  const string &clientFile, const string &assignFile)
{
    string delId;
    cout << "Enter Client ID to delete: ";
    getline(cin, delId);

    auto clientIt = find_if(clients.begin(), clients.end(), [&](const Client &c)
                            { return c.getId() == delId; });

    if (clientIt == clients.end())
    {
        cout << red("❌ No client found with that ID.") << endl;
        return;
    }

    cout << "Client found:\n";
    Table info;
    info.add_row({"ID", "Name", "Contact", "Company", "Address", "Service"});
    info.add_row({clientIt->getId(), clientIt->getName(), clientIt->getContact(),
                  clientIt->getCompany(), clientIt->getAddress(), clientIt->getService()});
    info[0].format().font_style({FontStyle::bold});
    cout << info << endl;

    char confirm;
    cout << "Are you sure you want to delete this client and all related employee? (y/n): ";
    cin >> confirm;
    cin.ignore();

    if (tolower(confirm) != 'y')
    {
        cout << red("❌ Deletion cancelled.") << endl;
        return;
    }
    clients.erase(clientIt);
    employee.erase(remove_if(employee.begin(), employee.end(), [&](const Employee &a)
                             { return a.clientId == delId; }),
                   employee.end());
    writeClientsToExcel(clientFile, clients);
    writeemployeeToExcel(assignFile, employee);
    cout << green("✅ Client and related employee deleted successfully!\n");
}

#endif