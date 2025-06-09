#include <iostream>
#include <tabulate/table.hpp>
#include <xlnt/xlnt.hpp>
using namespace std;
using namespace tabulate;

class Client
{
private:
    string id, name, contact, company, address, service;

public:
    Client(string id, string name, string contact, string company, string address, string service)
    {
        this->id = id;
        this->name = name;
        this->contact = contact;
        this->company = company;
        this->address = address;
        this->service = service;
    }
    string getId() const { return id; }
    string getName() const { return name; }
    string getContact() const { return contact; }
    string getCompany() const { return company; }
    string getAddress() const { return address; }
    string getService() const { return service; }
};

class Assignment
{
public:
    string clientId, employeeName;
    Assignment(string id, string emp)
    {
        clientId = id;
        employeeName = emp;
    }
};

void printClientTable(const vector<Client> &clients)
{
    Table table;
    table.add_row({"ID", "Name", "Contact", "Company", "Address", "Service"});
    for (auto &c : clients)
    {
        table.add_row({c.getId(), c.getName(), c.getContact(), c.getCompany(), c.getAddress(), c.getService()});
    }
    table[0].format().font_style({FontStyle::bold});
    cout << table << endl;
}

void printAssignmentTable(const vector<Assignment> &assignments)
{
    Table table;
    table.add_row({"Client ID", "Employee"});
    for (auto &a : assignments)
    {
        table.add_row({a.clientId, a.employeeName});
    }
    table[0].format().font_style({FontStyle::bold});
    cout << table << endl;
}

vector<Client> readClientsFromExcel(const string &filename)
{
    vector<Client> clients;
    xlnt::workbook wb;
    try
    {
        wb.load(filename);
    }
    catch (...)
    {
        cout << "⚠️ Failed to open clients Excel file.\n";
        return clients;
    }
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

vector<Assignment> readAssignmentsFromExcel(const string &filename)
{
    vector<Assignment> assignments;
    xlnt::workbook wb;
    try
    {
        wb.load(filename);
    }
    catch (...)
    {
        cout << "⚠️ Failed to open assignments Excel file.\n";
        return assignments;
    }
    auto ws = wb.active_sheet();
    for (auto row : ws.rows(false))
    {
        if (row[0].to_string() == "Client ID")
            continue;
        assignments.emplace_back(row[0].to_string(), row[1].to_string());
    }
    return assignments;
}

void writeAssignmentsToExcel(const string &filename, const vector<Assignment> &assignments)
{
    xlnt::workbook wb;
    auto ws = wb.active_sheet();
    ws.title("Assignments");
    ws.cell("A1").value("Client ID");
    ws.cell("B1").value("Employee");
    int row = 2;
    for (auto &a : assignments)
    {
        ws.cell("A" + to_string(row)).value(a.clientId);
        ws.cell("B" + to_string(row)).value(a.employeeName);
        row++;
    }
    wb.save(filename);
}
void deleteClient(vector<Client>& clients, vector<Assignment>& assignments,
                  const string& clientFile, const string& assignFile) {
    string delId;
    cout << "Enter Client ID to delete: ";
    getline(cin, delId);

    auto clientIt = std::find_if(clients.begin(), clients.end(), [&](const Client& c) {
        return c.getId() == delId;
    });

    if (clientIt == clients.end()) {
        cout << "❌ No client found with that ID.\n";
        return;
    }

    // Display client info before deletion
    cout << "Client found:\n";
    Table info;
    info.add_row({"ID", "Name", "Contact", "Company", "Address", "Service"});
    info.add_row({clientIt->getId(), clientIt->getName(), clientIt->getContact(),
                  clientIt->getCompany(), clientIt->getAddress(), clientIt->getService()});
    info[0].format().font_style({FontStyle::bold});
    cout << info << endl;

    // Ask for confirmation
    char confirm;
    cout << "Are you sure you want to delete this client and all related assignments? (y/n): ";
    cin >> confirm;
    cin.ignore();  // flush newline

    if (tolower(confirm) != 'y') {
        cout << "❌ Deletion cancelled.\n";
        return;
    }

    // Perform deletion
    clients.erase(clientIt);
    assignments.erase(
        std::remove_if(assignments.begin(), assignments.end(), [&](const Assignment& a) {
            return a.clientId == delId;
        }),
        assignments.end()
    );

    writeClientsToExcel(clientFile, clients);
    writeAssignmentsToExcel(assignFile, assignments);
    cout << "✅ Client and related assignments deleted successfully!\n";
}

int main()
{
    string clientFile = "clients.xlsx";
    string assignFile = "assignments.xlsx";

    vector<string> menu = {
        "Add Client Record",
        "Assign Employee to Client",
        "Show Clients",
        "Show Employee",
        "Delete Client",
        "Search Client by ID or Name",
        "Exit"};

    vector<Client> clients = readClientsFromExcel(clientFile);
    vector<Assignment> assignments = readAssignmentsFromExcel(assignFile);

    int option;
    do
    {
        Table t;
        t.add_row({"No", "Menu"});
        for (int i = 0; i < menu.size(); i++)
        {
            t.add_row({to_string(i + 1), menu[i]});
        }
        t[0].format().font_style({FontStyle::bold});
        cout << t << endl;
        cout << "Enter choice: ";
        cin >> option;
        cin.ignore();

        switch (option)
        {
        case 1:
        {
            string id, name, contact, company, address, service;
            cout << "ID: ";
            getline(cin, id);
            cout << "Name: ";
            getline(cin, name);
            cout << "Contact: ";
            getline(cin, contact);
            cout << "Company: ";
            getline(cin, company);
            cout << "Address: ";
            getline(cin, address);
            cout << "Service: ";
            getline(cin, service);
            clients.emplace_back(id, name, contact, company, address, service);
            writeClientsToExcel(clientFile, clients);
            cout << "✅ Client added successfully!\n";
        }
        break;
        case 2:
        {
            string clientId, empName;
            cout << "Client ID: ";
            getline(cin, clientId);
            cout << "Employee Name: ";
            getline(cin, empName);
            assignments.emplace_back(clientId, empName);
            writeAssignmentsToExcel(assignFile, assignments);
            cout << "✅ Employee assigned successfully!\n";
        }
        break;
        case 3:
            printClientTable(clients);
            break;
        case 4:
            printAssignmentTable(assignments);
            break;
        case 5:
            deleteClient(clients, assignments, clientFile, assignFile);
            break;
        case 6:
            string query;
            cout << "Enter Client ID or Name to search: ";
            getline(cin, query);
            bool found = false;
            for (const auto &c : clients)
            {
                if (c.getId() == query || c.getName() == query)
                {
                    Table result;
                    result.add_row({"ID", "Name", "Contact", "Company", "Address", "Service"});
                    result.add_row({c.getId(), c.getName(), c.getContact(), c.getCompany(), c.getAddress(), c.getService()});
                    result[0].format().font_style({FontStyle::bold});
                    cout << result << endl;
                    found = true;
                    break;
                }
            }
            if (!found)
                cout << "❌ No client found with that ID or Name.\n";
        break;
        }
    } while (option != 7);

    return 0;
}
