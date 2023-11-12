namespace Task3
{
    public interface IWare
    {
        int Code { get; set; }
        string Caption { get; set; }
        string Unit { get; set; }
        decimal Price { get; set; }
    }
}
