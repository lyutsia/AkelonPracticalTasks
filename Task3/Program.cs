namespace Task3
{
    class Program
    {
        static void Main(string[] args)
        {
            var exit = false;
            var fileLoaded = false;
            DataFromExcel dataFromExcel = new DataFromExcel();
            while (!exit)
            {
                Console.WriteLine("Выберите действие:\n"
                + "\t1-Открыть xlsx-файл;\r\n"
                + "\t2-По наименованию товара выводить информацию о клиентах, заказавших этот товар;\r\n"
                + "\t3-Запрос на изменение контактного лица клиента с указанием параметров: Название организации, ФИО нового контактного лица;\r\n"
                + "\t4-Запрос на определение золотого клиента, клиента с наибольшим количеством заказов, за указанный год, месяц.\r\n"
                + "\t5-Закрыть программу.\r\n");
                int.TryParse(Console.ReadLine(), out int numberAction);
                if (numberAction == 0)
                {
                    Console.WriteLine("Введено неверное значение!");
                    continue;
                }
                else
                {
                    switch (numberAction)
                    {
                        case 1:
                            {
                                Console.WriteLine("Введите путь до файла: ");
                                var pathFile = Console.ReadLine();
                                dataFromExcel = new DataFromExcel(pathFile);
                                fileLoaded = dataFromExcel.LoadDataFromExcel(pathFile);
                                break;
                            }
                        case 2:
                            {
                                if (!CheckFileLoaded(fileLoaded))
                                    continue;

                                Console.WriteLine("Введите название товара: ");
                                var caption = Console.ReadLine();
                                var text = dataFromExcel.GetInfoAboutClientsOrderedWareByCaption(caption);
                                Console.WriteLine(text);
                                break;
                            }
                        case 3:
                            {
                                if (!CheckFileLoaded(fileLoaded))
                                    continue;

                                Console.WriteLine("Введите название компании: ");
                                var companyName = Console.ReadLine();
                                Console.WriteLine("Введите ФИО нового контактного лица: ");
                                var newContactPersonName = Console.ReadLine();
                                var result = dataFromExcel.SetContactPersonNameByCompany(companyName, newContactPersonName);
                                if (result)
                                    Console.WriteLine("Данные успешно изменены!");
                                else
                                    Console.WriteLine("Произошла ошибка! Данные не были изменены");
                                break;
                            }
                        case 4:
                            {
                                if (!CheckFileLoaded(fileLoaded))
                                    continue;

                                Console.WriteLine("Вывести золотого клиента\n"
                                + "\t1-за месяц\n"
                                + "\t2-за год?");
                                int.TryParse(Console.ReadLine(), out int commandNumber);

                                if (commandNumber == 1)
                                {
                                    Console.WriteLine("Введите месяц в формате мм.гггг:");
                                    var dateSplit = Console.ReadLine().Split('.');
                                    if (dateSplit.Count() != 2)
                                    {
                                        Console.WriteLine("Неверный формат даты!");
                                        continue;
                                    }
                                    int.TryParse(dateSplit[0], out int month);
                                    int.TryParse(dateSplit[1], out int year);
                                    if (month > 12 || month < 1)
                                    {
                                        Console.WriteLine("Неверное значение месяца!");
                                        continue;
                                    }
                                    if (year > DateTime.Today.Year || year < 1000)
                                    {
                                        Console.WriteLine("Неверное значение года!");
                                        continue;
                                    }

                                    var dateTimeStart = new DateTime(year, month, 1);
                                    var goldClientInMonth = dataFromExcel.GetClientWithMaxOrdersInMonth(dateTimeStart);
                                    if (goldClientInMonth == null)
                                        Console.WriteLine("За этот месяц не было заказов");
                                    else
                                        Console.WriteLine($"Золотой клиент месяца: {goldClientInMonth.ContactPersonName}");
                                }
                                else if (commandNumber == 2)
                                {
                                    Console.WriteLine("Введите год:");
                                    var dateText = Console.ReadLine();
                                    int.TryParse(dateText, out int year);
                                    if (year > DateTime.Today.Year || year < 1000)
                                    {
                                        Console.WriteLine("Неверное значение года!");
                                        continue;
                                    }

                                    var dateTimeStart = new DateTime(year, 1, 1);
                                    var goldClientInYear = dataFromExcel.GetClientWithMaxOrdersInYear(dateTimeStart);
                                    if (goldClientInYear == null)
                                        Console.WriteLine("За этот год не было заказов");
                                    else
                                        Console.WriteLine($"Золотой клиент года: {goldClientInYear.ContactPersonName}");
                                }
                                else
                                {
                                    Console.WriteLine("Введено неверное значение!");
                                    continue;
                                }
                                break;
                            }
                        case 5:
                            {
                                exit = true;
                                break;
                            }
                        default:
                            {
                                Console.WriteLine("Введен неверный номер!");
                                break;
                            }
                    }
                }
            }
            Console.ReadKey();
        }

        static bool CheckFileLoaded(bool fileLoaded)
        {
            if (fileLoaded)
                return true;

            Console.WriteLine("Файл не загружен. Выберите действие 1");
            return false;
        }
    }
}
