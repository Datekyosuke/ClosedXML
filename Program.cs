using ClosedXML.Excel;
namespace ClosedXMLTwo
{
    public class Program
    {
        public static void Main()
        {
            string pathToExcelFile;
            Console.WriteLine("Введите путь до файла с данными (формат xlsx):");
            pathToExcelFile = Console.ReadLine();

            string filePath = "C:\\Users\\DateKyosuke.KOYRAVA\\Downloads\\Практическое задание для кандидата.xlsx";
            var wb = new XLWorkbook(filePath);
            var wsProducts = wb.Worksheet(1);
            var wsClients = wb.Worksheet(2);
            var wsOrders = wb.Worksheet(3);

            while (true)
            {
                Console.WriteLine("Введите номер запроса (1-4) или q для выхода: ");
                string choice = Console.ReadLine();

                if (choice == "1")
                {
                    Console.WriteLine("Введите наименование товара: ");
                    string productName = Console.ReadLine();
                    SearchClientsByProduct(productName, wsProducts, wsOrders, wsClients);
                }
                else if (choice == "2")
                {
                    Console.WriteLine("Введите название организации:");
                    string clientName = Console.ReadLine();
                    Console.WriteLine("Введите ФИО нового контактного лица:");
                    string representativeName = Console.ReadLine();
                    ChangeContactPerson(clientName, representativeName, wsClients);
                    wb.Save();
                }
                else if (choice == "3")
                {
                    Console.WriteLine("Введите год:");
                    var tryYear = int.TryParse(Console.ReadLine(), out var year);
                    Console.WriteLine("Введите месяц:");
                    var tryMonth = int.TryParse(Console.ReadLine(), out var month);
                    FindGoldenClient(year, month, wsOrders, wsClients);
                }
                else if (choice == "q") { break; }
                else
                {
                    Console.WriteLine("Неверный ввод.");
                }
            }


        }
        public static void SearchClientsByProduct(string productName, IXLWorksheet wsProducts, IXLWorksheet wsOrders, IXLWorksheet wsClients)
        {
            var productRow = wsProducts.RowsUsed()
            .FirstOrDefault(r => r.Cell("B").GetString() == productName);

            if (productRow == null)
            {
                Console.WriteLine("Товар не найден.");
                return;
            }

            var productId = int.Parse(productRow.Cell("A").GetString());
            var priceString = productRow.Cell("D").GetString();
            decimal price = decimal.Parse(priceString);
            var orders = wsOrders.RowsUsed().Skip(1)
                .Where(r => int.Parse(r.Cell("B").GetString()) == productId)
                .Select(r => new
                {
                    ClientId = int.Parse(r.Cell("C").GetString()),
                    Quantity = int.Parse(r.Cell("E").GetString()),
                    Date = r.Cell("F").GetString()
                });

            foreach (var order in orders)
            {
                var clientRow = wsClients.RowsUsed().Skip(1)
                    .FirstOrDefault(r => int.Parse(r.Cell("A").GetString()) == order.ClientId);
                if (clientRow != null)
                {
                    Console.WriteLine($"Клиент: {clientRow.Cell("B").GetString()}");
                    Console.WriteLine($"Количество товара: {order.Quantity}");
                    Console.WriteLine($"Цена товара: {price}");
                    Console.WriteLine($"Дата заказа: {order.Date}");
                    Console.WriteLine();
                }
            }
        }

        public static void ChangeContactPerson(string companyName, string newContactPerson, IXLWorksheet wsClients)
        {
            var clientRow = wsClients.RowsUsed()
                .FirstOrDefault(r => r.Cell("B").GetString() == companyName);

            if (clientRow == null)
            {
                Console.WriteLine("Клиент не найден.");
                return;
            }

            clientRow.Cell("D").Value = newContactPerson;
            Console.WriteLine("Контактное лицо успешно изменено.");

        }
        public static void FindGoldenClient(int year, int month, IXLWorksheet wsOrders, IXLWorksheet wsClients)
        {
            var ordersInPeriod = wsOrders.RowsUsed().Skip(1)
                .Where(r => (r.Cell("F").GetDateTime()).Year == year &&
                            (r.Cell("F").GetDateTime()).Month == month);

            var clientOrders = ordersInPeriod
                .GroupBy(r => int.Parse(r.Cell("C").GetString()))
                .Select(g => new
                {
                    ClientId = g.Key,
                    OrderCount = g.Count()
                });

            var goldenClient = clientOrders.OrderByDescending(c => c.OrderCount).FirstOrDefault();

            if (goldenClient != null)
            {
                var goldenClientRow = wsClients.RowsUsed().Skip(1)
                    .FirstOrDefault(r => int.Parse(r.Cell("A").GetString()) == goldenClient.ClientId);
                if (goldenClientRow != null)
                {
                    Console.WriteLine($"Золотой клиент: {goldenClientRow.Cell("B").GetString()} \n");
                }
            }
        }
    }
}