namespace Task5
{
    public class VacationDistribution
    {
        private const int VacationDays = 28;
        private const int DayCountBetweenVacationDifferentWorkers = 3;
        private const int MonthCountBetweenVacationWorker = 1;

        public Dictionary<string, List<Vacation>> GetVacationDistribution(List<string> workersName)
        {
            var vacationDictionary = new Dictionary<string, List<Vacation>>();
            foreach (var workerName in workersName)
            {
                vacationDictionary.Add(workerName, new List<Vacation>());
            }

            // Список отпусков сотрудников
            var allVacationList = new List<Vacation>();
            var gen = new Random();
            var dayOfYear = DateTime.Today.DayOfYear;
            var start = new DateTime(DateTime.Today.Year, 1, 1);
            var aviableWorkingDaysOfWeekWithoutWeekends = new List<DayOfWeek>()
            {
                DayOfWeek.Monday,
                DayOfWeek.Tuesday,
                DayOfWeek.Wednesday,
                DayOfWeek.Thursday,
                DayOfWeek.Friday
            };

            foreach (var vacationList in vacationDictionary.Values)
            {
                var vacationCount = 0;
                while (vacationCount < VacationDays)
                {
                    var startDate = start.AddDays(gen.Next(dayOfYear));

                    if (aviableWorkingDaysOfWeekWithoutWeekends.Contains(startDate.DayOfWeek))
                    {
                        var vactionStep = GetVacationStep(vacationCount, gen);
                        var endDate = startDate.AddDays(vactionStep);
                        if (CheckConditionsVacation(startDate, endDate, vacationList, allVacationList))
                        {
                            vacationCount += vactionStep;
                            var vacation = new Vacation() { StartDate = startDate, EndDate = endDate.AddDays(-1) };
                            vacationList.Add(vacation);
                            allVacationList.Add(vacation);
                        }

                    }
                }
            }
            return vacationDictionary;
        }

        /// <summary>
        /// получить продолжительность отпуска
        /// </summary>
        /// <param name="vacationCount"></param>
        /// <param name="gen"></param>
        /// <returns></returns>
        public int GetVacationStep(int vacationCount, Random gen)
        {
            int[] vacationSteps = { 7, 14 };
            int vacationIndex = gen.Next(vacationSteps.Length);
            var vactionStep = vacationSteps[vacationIndex];
            if (vacationCount + vactionStep > VacationDays)
                vactionStep = vacationSteps.Min();
            return vactionStep;
        }
        /// <summary>
        /// Проверка условий по отпуску
        /// </summary>
        /// <returns></returns>
        public bool CheckConditionsVacation(DateTime startDate, DateTime endDate, List<Vacation> vacationList, List<Vacation> allVacationList)
            => CheckConditionsForAllVacationList(startDate, endDate, allVacationList) && CheckConditionsInVacationList(startDate, endDate, vacationList);

        /// <summary>
        /// проверка условий отпусков одного сотрудника
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <param name="vacationList"></param>
        /// <returns></returns>
        public bool CheckConditionsInVacationList(DateTime startDate, DateTime endDate, List<Vacation> vacationList)
        {
            return !vacationList.Any(element => element.EndDate < startDate && element.EndDate.AddMonths(MonthCountBetweenVacationWorker) >= startDate)
                      && !vacationList.Any(element => element.StartDate > endDate && element.StartDate.AddMonths(-MonthCountBetweenVacationWorker) <= endDate);
        }
        /// <summary>
        /// проверка условий отпусков разных сотрудников
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <param name="allVacationList"></param>
        /// <returns></returns>
        public bool CheckConditionsForAllVacationList(DateTime startDate, DateTime endDate, List<Vacation> allVacationList)
        {
            //проверка пересекаются ли отпуска
            return !allVacationList.Any(element => IsIncludedInterval(element.StartDate, startDate, endDate)
                || IsIncludedInterval(startDate, element.StartDate, element.EndDate))

                // есть ли промежуток между отпусками разных сотудников
                && !allVacationList.Any(element => IsIncludedInterval(element.EndDate.AddDays(DayCountBetweenVacationDifferentWorkers), startDate, endDate))
                && !allVacationList.Any(element => IsIncludedInterval(element.StartDate.AddDays(-DayCountBetweenVacationDifferentWorkers), startDate, endDate));
        }

        public bool IsIncludedInterval(DateTime verifiableDate, DateTime startDate, DateTime endDate)
            => verifiableDate >= startDate && verifiableDate <= endDate;
    }

    public struct Vacation
    {
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }

    }
}
