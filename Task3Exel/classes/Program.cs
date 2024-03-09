namespace Task3Exel;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("Введите путь до файла: ");
        string? filePath = Console.ReadLine();

        if (!File.Exists(filePath)) //C:\Users\Vitaly\source\repos\Task3Exel\Task3Exel\EnteringData.xlsx
        {
            Console.WriteLine("Файл не существует!");
            return;
        }

        if (File.ReadAllText(filePath).Length == 0)
        {
            Console.WriteLine("Файл пуст!");
            return;
        }

        List<Product> products = LoadDataFromFile.LoadProducts(filePath);
        List<Client> clients = LoadDataFromFile.LoadClients(filePath);
        List<Request> requests = LoadDataFromFile.LoadRequests(filePath);

        var userInterface = new UserInterface();
        userInterface.Run(clients, products, requests, filePath);
    }
}
