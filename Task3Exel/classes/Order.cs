namespace Task3Exel;

class Order
{
    public string CustomerName { get; }
    public string OrganizationName { get; }
    public int Quantity { get; }
    public decimal Price { get; }
    public DateTime OrderDate { get; }
    public string ContactPerson { get; set; }
}
