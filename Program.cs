using ClosedXML.Excel;
namespace ClosedXMLTwo
{
    ///<summary>
    ///Конслоьное приложение выполняет запросы к данным в файле Excel,  
    ///ищет клиентов по товару, может изменить данные контактного лица и умеет находить золотого клиента
    ///</summary>
    public class Program
    {
        public static void Main()
        {
            // Получение пути к файлу с данными
            string pathToExcelFile;
            Console.WriteLine("Введите путь до файла с данными (формат xlsx):");
            pathToExcelFile = Console.ReadLine();

            // Открытие рабочей книги и получение листов с данными
            var wb = new XLWorkbook(pathToExcelFile);
            var wsProducts = wb.Worksheet(1);
            var wsClients = wb.Worksheet(2);
            var wsOrders = wb.Worksheet(3);

            // Цикл для обработки пользовательских запросов
            while (true)
            {

                Console.WriteLine("Введите номер запроса (1-4) или q для выхода: ");
                string choice = Console.ReadLine();

                // Обработка выбора пользователя

                // Запрос 1: поиск клиентов по товару
                if (choice == "1")
                {
                    Console.WriteLine("Введите наименование товара: ");
                    string productName = Console.ReadLine();
                    SearchClientsByProduct(productName, wsProducts, wsOrders, wsClients);
                }

                // Запрос 2: изменение контактного лица клиента
                else if (choice == "2")
                {
                    Console.WriteLine("Введите название организации:");
                    string clientName = Console.ReadLine();
                    Console.WriteLine("Введите ФИО нового контактного лица:");
                    string representativeName = Console.ReadLine();
                    ChangeContactPerson(clientName, representativeName, wsClients);
                    wb.Save();
                }

                // Запрос 3: поиск золотого клиента
                else if (choice == "3")
                {
                    Console.WriteLine("Введите год:");
                    var tryYear = int.TryParse(Console.ReadLine(), out var year);
                    Console.WriteLine("Введите месяц:");
                    var tryMonth = int.TryParse(Console.ReadLine(), out var month);
                    FindGoldenClient(year, month, wsOrders, wsClients);
                }
                // Выход из программы
                else if (choice == "q") { break; }
                else
                {
                    Console.WriteLine("Неверный ввод.");
                }
            }


        }
        // Поиск клиентов по товару
        public static void SearchClientsByProduct(string productName, IXLWorksheet wsProducts, IXLWorksheet wsOrders, IXLWorksheet wsClients)
        {
            // Поиск строки товара по наименованию
            var productRow = wsProducts.RowsUsed()
            .FirstOrDefault(r => r.Cell("B").GetString() == productName);

            if (productRow == null)
            {
                Console.WriteLine("Товар не найден.");
                return;
            }

            // Получение идентификатора товара и цены
            var productId = int.Parse(productRow.Cell("A").GetString());
            var priceString = productRow.Cell("D").GetString();
            decimal price = decimal.Parse(priceString);

            // Получение заказов по товару
            var orders = wsOrders.RowsUsed().Skip(1)
                .Where(r => int.Parse(r.Cell("B").GetString()) == productId)
                .Select(r => new
                {
                    ClientId = int.Parse(r.Cell("C").GetString()),
                    Quantity = int.Parse(r.Cell("E").GetString()),
                    Date = r.Cell("F").GetString()
                });

            // Вывод информации о клиентах, заказавших товар
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

        // Изменение контактного лица клиента
        public static void ChangeContactPerson(string companyName, string newContactPerson, IXLWorksheet wsClients)
        {
            // Поиск строки клиента по названию организации
            var clientRow = wsClients.RowsUsed()
                .FirstOrDefault(r => r.Cell("B").GetString() == companyName);

            if (clientRow == null)
            {
                Console.WriteLine("Клиент не найден.");
                return;
            }

            // Изменение значения ячейки с контактным лицом
            clientRow.Cell("D").Value = newContactPerson;
            Console.WriteLine("Контактное лицо успешно изменено.");

        }

        // Поиск золотого клиента
        public static void FindGoldenClient(int year, int month, IXLWorksheet wsOrders, IXLWorksheet wsClients)
        {
            // Получение заказов за указанный период
            var ordersInPeriod = wsOrders.RowsUsed().Skip(1)
                .Where(r => (r.Cell("F").GetDateTime()).Year == year &&
                            (r.Cell("F").GetDateTime()).Month == month);

            // Группировка заказов по клиентам и подсчет количества заказов для каждого клиента
            var clientOrders = ordersInPeriod
                .GroupBy(r => int.Parse(r.Cell("C").GetString()))
                .Select(g => new
                {
                    ClientId = g.Key,
                    OrderCount = g.Count()
                });

            // Получение золотого клиента с максимальным количеством заказов
            var goldenClient = clientOrders.OrderByDescending(c => c.OrderCount).FirstOrDefault();

            //Если клиент найден, то вывести его на экран
            if (goldenClient != null)
            {
                var goldenClientRow = wsClients.RowsUsed().Skip(1)
                    .FirstOrDefault(r => int.Parse(r.Cell("A").GetString()) == goldenClient.ClientId);
                    Console.WriteLine($"Золотой клиент: {goldenClientRow.Cell("B").GetString()} \n");

            }
        }
    }
}