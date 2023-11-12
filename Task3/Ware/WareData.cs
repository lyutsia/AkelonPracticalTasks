namespace Task3
{
    public class WareData : IData
    {
        private List<Ware> _wares;
        public WareData()
        {
            _wares = new List<Ware>();
        }


        public void AddObject(object obj)
        {
            if (obj is not Ware ware)
                return;
            _wares.Add(ware);
        }
        public string SheetName => "Товары";
        public Type ObjectType => typeof(Ware);
        public Dictionary<string, string> Headers { get; set; } = new Dictionary<string, string>();
        public object GetObjectByCode(int code) => _wares.FirstOrDefault(w => w.Code == code);
        public IEnumerable<Ware> GetWareByCaption(string caption) => _wares.Where(w => w.Caption == caption);
    }
}
