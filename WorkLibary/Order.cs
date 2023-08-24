namespace WorkLibary;

public class Order
{
    public readonly OrderTime OrderTime;
    public  Lunch.Lunch? Lunch;
    public HotFood? HotFood;
    public Soup? Soup;
    public Bakery? Bakery;
    public readonly bool WillCoffee;
    public readonly Days Day;

    public Order(OrderTime orderTime, Lunch.Lunch? lunch, HotFood? hotFood, Soup? soup, Bakery? bakery, bool willCoffee, Days day)
    {
        OrderTime = orderTime;
        Lunch = lunch;
        HotFood = hotFood;
        Soup = soup;
        Bakery = bakery;
        WillCoffee = willCoffee;
        Day = day;
    }
}