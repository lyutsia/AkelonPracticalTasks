namespace Task3
{
    public class Ware : IWare
    {
        public Ware()
        {

        }
        public Ware(int code)
        {
            Code = code;
        }
        public Ware(int code, string caption) : this(code)
        {
            Caption = caption;
        }

        [ColumnCaptionAttribute("Код товара", typeof(int))]
        public int Code { get; set; }

        [ColumnCaptionAttribute("Наименование", typeof(string))]
        public string Caption { get; set; }

        [ColumnCaptionAttribute("Ед. измерения", typeof(string))]
        public string Unit { get; set; }

        [ColumnCaptionAttribute("Цена товара за единицу", typeof(decimal))]
        public decimal Price { get; set; }
    }
}
