#include <iostream>
#include <tabulate/table.hpp>
#include <xlnt/xlnt.hpp>
#include <exception>
#include "function.h"

int main()
{
    system("cls");
    string clientFile = "clients.xlsx";
    string assignFile = "employee.xlsx";

    vector<string> menu = {
        "Add Client Record",
        "Assign Employee to Client",
        "Show Clients",
        "Show Employee",
        "Delete Client",
        "Search Client by ID or Name",
        "Exit"};

    vector<Client> clients = readClientsFromExcel(clientFile);
    vector<Employee> employee = reademployeeFromExcel(assignFile);

    int option;
    do
    {
        system("cls");
        Table t;
        cout << "\033[1;36m=== Client Management System ===\033[0m\n";
        t.add_row({"No", "Menu"});
        for (int i = 0; i < menu.size(); i++)
        {
            t.add_row({to_string(i + 1), menu[i]});
        }
        t[0].format().font_style({FontStyle::bold});
        cout << t << endl;
        cout << bold_blue("Enter choice: ");
        cin >> option;
        cin.ignore();

        switch (option)
        {
        case 1:
        {
            system("cls");
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
            cout << green("✅ Client added successfully!") << endl;
            break;
        }
        case 2:
        {
            system("cls");
            string clientId, empName;
            cout << "Client ID: ";
            getline(cin, clientId);
            cout << "Employee Name: ";
            getline(cin, empName);
            employee.emplace_back(clientId, empName);
            writeemployeeToExcel(assignFile, employee);
            cout << green("✅ Employee assigned successfully!") << endl;
            pressEnter();
            break;
        }
        case 3:
        {
            system("cls");
            printClientTable(clients);
            pressEnter();
            break;    
        }
        case 4:
            system("cls");
            printEmployeeTable(employee);
            pressEnter();
            break;
        case 5:
            system("cls");
            deleteClient(clients, employee, clientFile, assignFile);
           pressEnter();
            break;
        case 6:
        {
            system("cls");
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
                cout << "\033[1;31m❌ No client found with that ID or Name.\033[0m\n";
            pressEnter();
            break;
        }
        default:
            cout << "Program closed!!\n";
            break;
        }
    } while (option != 7);

    return 0;
}
