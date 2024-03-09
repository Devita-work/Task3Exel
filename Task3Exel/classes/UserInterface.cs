using ClosedXML.Excel;

namespace Task3Exel;

public class UserInterface
{
    public void Run(List<Client> clients, List<Product> products, List<Request> requests, string filePath)
    {
        while (true)
        {
            Console.Clear();
            Console.WriteLine("**Выберите действие:**");
            Console.WriteLine("1. Информация о клиентах по товару");
            Console.WriteLine("2. Изменение контактного лица клиента");
            Console.WriteLine("3. Определение золотого клиента");
            Console.WriteLine("4. Выход");

            var choice = Console.ReadLine();

            if (choice == null || choice.ToString() == "")
                continue;

            switch (int.Parse(choice))
            {
                case 1:
                    GetClientsByProduct(clients, products, requests);
                    break;
                case 2:
                    ChangeClientContactPerson(clients, filePath);
                    break;
                case 3:
                    GetGoldenClient(clients, requests);
                    break;
                case 4:
                    return;
                default:
                    Console.WriteLine("Неверный выбор!");
                    break;
            }
        }
    }

    private void GetClientsByProduct(List<Client> clients, List<Product> products, List<Request> requests)
    {
        Console.WriteLine("Введите наименование товара:");
        string productName = Console.ReadLine();

        var product = products.Find(p => p.Name.ToLower() == productName.ToLower());

        if (product == null)
        {
            Console.WriteLine("Товар не найден!");
            return;
        } 

        var requestsForProduct = requests.Where(r => r.ProductCode == product.ProductCode);

        foreach (var request in requestsForProduct)
        {
            var client = clients.Find(c => c.ClientCode == request.ClientCode);

            Console.WriteLine();
            Console.WriteLine($"**Информация о клиенте:**");
            Console.WriteLine($"Название организации: {client.OrganizationName}");
            Console.WriteLine($"Адрес: {client.Address}");
            Console.WriteLine($"Контактное лицо: {client.ContactPerson}");

            Console.WriteLine();
            Console.WriteLine($"**Информация о заказе:**");
            Console.WriteLine($"Количество: {request.RequiredQuantity}");
            Console.WriteLine($"Цена: {product.Price}");
            Console.WriteLine($"Дата заказа: {request.PlacementDate:d}");
        }

        Console.ReadKey();
    }

    private static void ChangeClientContactPerson(List<Client> clients, string filePath)
    {
        Console.WriteLine("Введите название организации:");
        string? organizationName = Console.ReadLine();

        Console.WriteLine("Введите ФИО нового контактного лица:");
        string? newContactPerson = Console.ReadLine();

        var client = clients.FirstOrDefault(c => c.OrganizationName.ToLower() == organizationName.ToLower());

        if (client == null)
        {
            Console.WriteLine("Клиент не найден!");
            return;
        }

        client.ContactPerson = newContactPerson;

        using (var workbook = new XLWorkbook(filePath))
        {
            var worksheet = workbook.Worksheet("Клиенты");
            var row = worksheet.Rows().FirstOrDefault(r => r.Cell(2).Value.ToString() == client.OrganizationName);

            string originalValue = row.Cell(4).Value.ToString();

            row.Cell(4).Value = newContactPerson;
            var updatedValue = row.Cell(4).Value.ToString();
            workbook.Save();

            if (originalValue != updatedValue)
            {
                Console.WriteLine("Контактное лицо успешно обновлено!");
            }
            else
            {
                Console.WriteLine("Ошибка при обновлении контактного лица!");
            }
        }
    }

    private static void GetGoldenClient(List<Client> clients, List<Request> requests)
    {
        Console.WriteLine("Введите год:");
        int year = int.Parse(Console.ReadLine());

        Console.WriteLine("Введите месяц:");
        int month = int.Parse(Console.ReadLine());

        var clientOrders = requests
            .Where(r => r.PlacementDate.Year == year && r.PlacementDate.Month == month)
            .GroupBy(r => r.ClientCode)
            .Select(g => new { ClientCode = g.Key, OrderCount = g.Count() })
            .OrderByDescending(g => g.OrderCount);

        var goldenCustomer = clientOrders.FirstOrDefault();

        if (goldenCustomer != null)
        {
            var client = clients.Find(c => c.ClientCode == goldenCustomer.ClientCode);

            Console.WriteLine();
            Console.WriteLine($"**Информация о золотом клиенте:**");
            Console.WriteLine($"Название организации: {client.OrganizationName}");
            Console.WriteLine($"Адрес: {client.Address}");
            Console.WriteLine($"Контактное лицо: {client.ContactPerson}");
            Console.WriteLine($"Количество заказов: {goldenCustomer.OrderCount}");
        }
        else
        {
            Console.WriteLine($"Золотой клиент за {GetMonthName(month)} {year} не найден.");
        }

        Console.ReadKey();
    }

    // Вспомогательный метод для получения названия месяца
    static string GetMonthName(int month)
    {
        return new DateTime(2022, month, 1).ToString("MMMM");
    }
}