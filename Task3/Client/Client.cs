namespace Task3
{
    public class Client : IClient
    {
        public Client()
        {
        }
        public Client(int code)
        {
            Code = code;
        }
        public Client(int code, string contactPersonName) : this(code)
        {
            ContactPersonName = contactPersonName;
        }
        public Client(int code, string contactPersonName, string companyName) : this(code, contactPersonName)
        {
            CompanyName = companyName;
        }

        [ColumnCaptionAttribute("Код клиента", typeof(int))]
        public int Code { get; set; }

        [ColumnCaptionAttribute("Наименование организации", typeof(string))]
        public string CompanyName { get; set; }

        [ColumnCaptionAttribute("Адрес", typeof(string))]
        public string Address { get; set; }

        [ColumnCaptionAttribute("Контактное лицо (ФИО)", typeof(string))]
        public string ContactPersonName { get; set; }

        public override string ToString() => $"{ContactPersonName} ({CompanyName})";
    }
}
