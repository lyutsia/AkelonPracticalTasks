namespace Task3
{
    public class Order : IOrder
    {
        public Order()
        {

        }
        public Order(int code)
        {
            Code = code;
        }

        [ColumnCaptionAttribute("Код заявки", typeof(int))]
        public int Code { get; set; }

        [ColumnCaptionAttribute("Номер заявки", typeof(int))]
        public int Number { get; set; }

        [ColumnCaptionAttribute("Дата размещения", typeof(DateTime))]
        public DateTime DatePosting { get; set; }

        [ColumnCaptionAttribute("Требуемое количество", typeof(int))]
        public int RequiredQuantity { get; set; }

        [ColumnCaptionAttribute("Код товара", typeof(Ware))]
        public IWare Ware { get; set; }

        [ColumnCaptionAttribute("Код клиента", typeof(Client))]
        public IClient Client { get; set; }
    }
}
