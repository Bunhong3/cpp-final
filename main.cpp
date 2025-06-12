#include "function.h"
int main()
{
    system("cls");
    string clientFile = "clients.xlsx";
    string empFile = "employee.xlsx";

    vector<string> menu = {
        "Add Client Record",
        "Assign Employee to Client",
        "Track Client ",
        "View All Clients",
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
            string name, contact, email;
            cout << "Enter Client Name    : ";
            getline(cin, name);
            cout << "Enter Contact Persion: ";
            getline(cin, contact);
            cout << "Enter Contact Email  : ";
            getline(cin, email);

            clients.emplace_back(name, contact, email);
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
            system("cls");
            printHeader("Show Employee");
            printEmployeeTable(employee);
            pressEnter();
            break;
        case 4:
        {
            system("cls");
            printHeader("Client Lists");
            printClientTable(clients);
            pressEnter();
            break;
        }
        
        default:
            cout << red("Program closed!!\n");
            break;
        }
    } while (option != 5);

    return 0;
}
