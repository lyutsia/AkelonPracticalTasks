namespace Task5
{
    class Program
    {
        static void Main(string[] args)
        {
            var workersName = new List<string>()
            {
                "Иванов Иван Иванович",
                "Петров Петр Петрович",
                "Юлина Юлия Юлиановна",
                "Сидоров Сидор Сидорович",
                "Павлов Павел Павлович",
                "Георгиев Георг Георгиевич",
            };
            var vacationDictionary = new VacationDistribution().GetVacationDistribution(workersName);
            PrintVacationDays(vacationDictionary);
            Console.ReadKey();
        }

        /// <summary>
        /// вывод отпусков
        /// </summary>
        /// <param name="vacationDictionary"></param>
        static void PrintVacationDays(Dictionary<string, List<Vacation>> vacationDictionary)
        {
            foreach (var vacationDictionaryItem in vacationDictionary)
            {
                Console.WriteLine("Дни отпуска " + vacationDictionaryItem.Key + " : ");
                foreach (var vacation in vacationDictionaryItem.Value)
                {
                    Console.WriteLine($"{vacation.StartDate.ToShortDateString()}-{vacation.EndDate.ToShortDateString()}");
                }
                Console.WriteLine();
            }
        }
    }
}
