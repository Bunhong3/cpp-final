
#include "function.h"
int main()
{
    system("cls");
    string clientFile = "clients.xlsx";
    string empFile = "employee.xlsx";

    vector<string> menu = {
        "Add Client Record",
        "Assign Employee to Client",
        "Show Clients",
        "Show Employee",
        "Delete Client",
        "Search Client by ID or Name",
        "Exit"};

    vector<Client> clients = readClientsFromExcel(clientFile);
    vector<Employee> employee = reademployeeFromExcel(empFile);

    int option;
    do
    {
        system("cls");
        Table t;
        printAppLogo();
        t.add_row({"No", "Menu"});
        for (int i = 0; i < menu.size(); i++)
        {
            t.add_row({to_string(i + 1), menu[i]});
        }
        t[0].format().font_style({FontStyle::bold}).font_align(FontAlign::center);
        t[0].format().font_color(Color::yellow);
        for (int i = 1; i <= menu.size(); i++)
        {
            if (i == 7) // Exit option
                t[i][1].format().font_color(Color::red);
            else
                t[i][1].format().font_color(Color::cyan);
        }
        cout << t << endl;
        cout << bold_blue(">> Enter choice: ");
        cin >> option;
        cin.ignore();

        switch (option)
        {
        case 1:
        {
            system("cls");
            printtHeader("📋 Add Client Record");
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
            cout << endl;
            cout << green("✅ Client added successfully!") << endl;
            pressEnter();
            break;
        }
        case 2:
        {
            system("cls");
            printHeader("Add Employee to Client");
            string clientId, empName;
            cout << "Client ID: ";
            getline(cin, clientId);
            cout << "Employee Name: ";
            getline(cin, empName);
            employee.emplace_back(clientId, empName);
            writeemployeeToExcel(empFile, employee);
            cout << endl;
            cout << green("✅ Employee emped successfully!") << endl;
            pressEnter();
            break;
        }
        case 3:
        {
            system("cls");
            printHeader("Client Lists");
            printClientTable(clients);
            pressEnter();
            break;
        }
        case 4:
            system("cls");
            printHeader("Show Employee");
            printEmployeeTable(employee);
            pressEnter();
            break;
        case 5:
            system("cls");
            printHeader("Dellete Client");
            deleteClient(clients, employee, clientFile, empFile);
            pressEnter();
            break;
        case 6:
        {
            system("cls");
            printHeader("Search Client by ID or Name");

            string query;
            cout << bold_blue("Enter Client ID or Name to search: ");
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
                cout << "\033[1;31m❌ No client found with that ID or Name!!\033[0m\n";
            pressEnter();
            break;
        }
        default:
            cout <<red("Program closed!!\n");
            break;
        }
    } while (option != 7);

    return 0;
}
