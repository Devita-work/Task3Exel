using ClosedXML.Excel;

namespace Task3Exel;
public static class LoadDataFromFile
{
    public static List<Product> LoadProducts(string filePath)
    {
        List<Product> products = new List<Product>();

        using (var workbook = new XLWorkbook(filePath))
        {
            var worksheet = workbook.Worksheet("Товары");

            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                int productCode = row.Cell(1).GetValue<int>();
                string productName = row.Cell(2).GetValue<string>();
                string unit = row.Cell(3).GetValue<string>();
                decimal price = row.Cell(4).GetValue<decimal>();

                products.Add(new Product { ProductCode = productCode, Name = productName, Unit = unit, Price = price });
            }
        }

        return products;
    }

    public static List<Client> LoadClients(string filePath)
    {
        List<Client> clients = new List<Client>();

        using (var workbook = new XLWorkbook(filePath))
        {
            var worksheet = workbook.Worksheet("Клиенты");

            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                int clientCode = row.Cell(1).GetValue<int>();
                string organizationName = row.Cell(2).GetValue<string>();
                string address = row.Cell(3).GetValue<string>();
                string contactPerson = row.Cell(4).GetValue<string>();

                clients.Add(new Client { ClientCode = clientCode, OrganizationName = organizationName, Address = address, ContactPerson = contactPerson });
            }
        }

        return clients;
    }

    public static List<Request> LoadRequests(string filePath)
    {
        List<Request> requests = new List<Request>();

        using (var workbook = new XLWorkbook(filePath))
        {
            var worksheet = workbook.Worksheet("Заявки");

            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                int requestCode = row.Cell(1).GetValue<int>();
                int productCode = row.Cell(2).GetValue<int>();
                int clientCode = row.Cell(3).GetValue<int>();
                int orderNumber = row.Cell(4).GetValue<int>();
                int quantity = row.Cell(5).GetValue<int>();
                DateTime placementDate = row.Cell(6).GetValue<DateTime>();

                requests.Add(new Request
                {
                    RequestCode = requestCode,
                    ProductCode = productCode,
                    ClientCode = clientCode,
                    OrderNumber = orderNumber,
                    RequiredQuantity = quantity,
                    PlacementDate = placementDate
                });
            }
        }

        return requests;
    }
}