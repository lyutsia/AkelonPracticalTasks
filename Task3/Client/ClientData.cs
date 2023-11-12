namespace Task3
{
    public class ClientData : IData
    {
        private List<Client> _clients;
        public ClientData()
        {
            _clients = new List<Client>();
        }

        public void AddObject(object obj)
        {
            if (obj is not Client client)
                return;
            _clients.Add(client);
        }
        public string SheetName => "Клиенты";
        public Type ObjectType => typeof(Client);
        public Dictionary<string, string> Headers { get; set; } = new Dictionary<string, string>();
        public object GetObjectByCode(int code) => _clients.FirstOrDefault(c => c.Code == code);
        public Client GetClientByCompany(string company) => _clients.FirstOrDefault(с => с.CompanyName == company);
    }
}
