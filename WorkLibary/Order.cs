namespace WorkLibary;

public class Order
{
    public readonly OrderTime OrderTime;
    public readonly Lunch.Lunch? Lunch;
    public readonly HotFood? HotFood;
    public readonly Soup? Soup;
    public readonly Bakery? Bakery;
    public readonly bool WillCoffee;

    public Order(OrderTime orderTime, Lunch.Lunch? lunch, HotFood? hotFood, Soup? soup, Bakery? bakery, bool willCoffee)
    {
        OrderTime = orderTime;
        Lunch = lunch;
        HotFood = hotFood;
        Soup = soup;
        Bakery = bakery;
        WillCoffee = willCoffee;
    }
}