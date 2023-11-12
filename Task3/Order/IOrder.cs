namespace Task3
{
    public interface IOrder
    {
        int Number { get; set; }
        DateTime DatePosting { get; set; }
        int RequiredQuantity { get; set; }
        IWare Ware { get; set; }
        IClient Client { get; set; }
    }
}
