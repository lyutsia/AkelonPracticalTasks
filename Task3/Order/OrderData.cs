namespace Task3
{
    public class OrderData : IData
    {
        private List<Order> _orders;
        public OrderData()
        {
            _orders = new List<Order>();
        }

        public void AddObject(object obj)
        {
            if (obj is not Order order)
                return;
            _orders.Add(order);
        }
        public string SheetName => "Заявки";
        public Type ObjectType => typeof(Order);
        public Dictionary<string, string> Headers { get; set; } = new Dictionary<string, string>();

        public object GetObjectByCode(int code) => _orders.FirstOrDefault(o => o.Code == code);

        public IEnumerable<Order> GetOrdersWithFindWares(IEnumerable<Ware> findWares) => _orders.Where(o => findWares.Contains(o.Ware));

        public IClient GetClientWithMaxOrdersInMonth(DateTime dateTimeStart) => GetClientWithOrdersCount(dateTimeStart, dateTimeStart.AddMonths(1));

        public IClient GetClientWithMaxOrdersInYear(DateTime dateTimeStart) => GetClientWithOrdersCount(dateTimeStart, dateTimeStart.AddYears(1));

        public IClient GetClientWithOrdersCount(DateTime dateTimeStart, DateTime dateTimeEnd)
        {
            var clientWithOrdersCount = _orders.Where(o => o.DatePosting >= dateTimeStart && o.DatePosting < dateTimeEnd)
                                               .GroupBy(o => o.Client)
                                               .Select(o => new { Client = o.Key, Count = o.Count() });
            if (clientWithOrdersCount.Count() == 0)
                return null;

            var maxOrderCount = clientWithOrdersCount.Select(o => o.Count).Max();
            return clientWithOrdersCount.FirstOrDefault(o => o.Count == maxOrderCount).Client;
        }
    }
}
