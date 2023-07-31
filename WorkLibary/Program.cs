using OfficeOpenXml;

var fi = new FileInfo(@"A.xlsx");
var users = new List<User>();
using (var p = new ExcelPackage(fi))
{
    var ws = p.Workbook.Worksheets["Sheet"];
    var rows = 42;
    for (var row = 2; row < rows; row++)
    {
        var user = new User((string) ws.Cells[row, 3].Value);
        user.AddOrder(ws.Cells[row, 5].Value, ws.Cells[row, 6].Value, ws.Cells[row, 7].Value, ws.Cells[row, 8].Value,
            ws.Cells[row, 9].Value, ws.Cells[row, 10].Value);

        users.Add(user);
    }
    
    p.Save();
}

using (var p = new ExcelPackage())
{
    var column = 1;
    var row = 1;
    
    foreach (var day in new[] {Days.Tuesday, Days.Wednesday, Days.Thursday, Days.Friday})
    {
        var ws = p.Workbook.Worksheets.Add(day.ToString());
        foreach (var user in users)
        {
            ws.Cells[row, column++].Value = user.Name;
            ws.Cells[row, column++].Value = user.Orders[(int) day].OrderTime;
            ws.Cells[row, column++].Value = user.Orders[(int) day].Lunch;
            ws.Cells[row, column++].Value = user.Orders[(int) day].HotFood;
            ws.Cells[row, column++].Value = user.Orders[(int) day].Soup;
            ws.Cells[row, column++].Value = user.Orders[(int) day].Bakery;
            ws.Cells[row, column++].Value = user.Orders[(int) day].WillCoffee;
            
            // ws.Cells[row++, column].Value = user.Name;
            // ws.Cells[row++, column].Value = user.Orders[(int) day].OrderTime;
            // ws.Cells[row++, column].Value = user.Orders[(int) day].Lunch;
            // ws.Cells[row++, column].Value = user.Orders[(int) day].HotFood;
            // ws.Cells[row++, column].Value = user.Orders[(int) day].Soup;
            // ws.Cells[row++, column].Value = user.Orders[(int) day].Bakery;
            // ws.Cells[row++, column].Value = user.Orders[(int) day].WillCoffee;
        }

        row++;
        column = 1;
    }
    
    p.SaveAs(new FileInfo(@"AAA.xlsx"));
}

Console.WriteLine();

public enum Days
{
    Tuesday = 0,
    Wednesday = 1,
    Thursday = 2,
    Friday = 3
}

public enum OrderTime
{
    Morning,
    Day,
    Night,
    Default
}

public enum Lunch
{
    Manager1,
    Manager2, 
    BusinessLady1, 
    BusinessLady2, 
    Freelancer1, 
    Freelancer2, 
    Gamer1, 
    Gamer2, 
    Vegan1, 
    Vegan2, 
    Vegan3, 
    Prince1, 
    Prince2, 
}

public enum Soup
{
    SpinachSoup, 
    MushroomSoup, 
    PumpkinSoup, 
}

public enum HotFood
{
    Pork, 
    Beef, 
    Chicken, 
    Shrimp, 
    FalafelBeans, 
    FalafelChickpea, 
    FalafelBuckwheat, 
    KebabChicken, 
    KebabPork, 
}

public enum Bakery
{
    AppleStrudel, 
    CarrotCake, 
    ChocolateCroissant, 
    CottageCheesePie, 
    CottageCheeseAndCherryPie, 
    RoseWithApplesAndCherries, 
}

public class User
{
    public readonly string Name;

    public readonly List<Order> Orders = new();

    public User(string name)
    {
        Name = name;
    }

    public void AddOrder(object date, object lunch, object hotFood, object soup, object bakery, object willCoffee)
    {
        var time = GetOrderTime((string) date);
        var lunchReady = GetLunch((string) lunch);
        var hotFoodReady = GetHotFood((string) hotFood);
        var soupReady = GetSoup((string) soup);
        var bakeryReady = GetBakery((string) bakery);
        
        Orders.Add(new Order(time, lunchReady, hotFoodReady, soupReady, bakeryReady, ((string) willCoffee) == "Да"));
    }

    private Bakery? GetBakery(string bakery)
    {
        return bakery switch
        {
            StringConstants.AppleStrudel => Bakery.AppleStrudel,
            StringConstants.CarrotCake => Bakery.CarrotCake,
            StringConstants.ChocolateCroissant => Bakery.ChocolateCroissant,
            StringConstants.CottageCheesePie => Bakery.CottageCheesePie,
            StringConstants.CottageCheeseAndCherryPie => Bakery.CottageCheeseAndCherryPie,
            StringConstants.RoseWithApplesAndCherries => Bakery.RoseWithApplesAndCherries,
            _ => null,
        };
    }

    private Soup? GetSoup(string soup)
    {
        return soup switch
        {
            StringConstants.SpinachSoup => Soup.SpinachSoup,
            StringConstants.MushroomSoup => Soup.MushroomSoup,
            StringConstants.PumpkinSoup => Soup.PumpkinSoup,
            _ => null,
        };
    }

    private HotFood? GetHotFood(string hotFood)
    {
        return hotFood switch
        {
            StringConstants.Pork => HotFood.Pork,
            StringConstants.Beef => HotFood.Beef,
            StringConstants.Chicken => HotFood.Chicken,
            StringConstants.Shrimp => HotFood.Shrimp,
            StringConstants.FalafelBeans => HotFood.FalafelBeans,
            StringConstants.FalafelChickpea => HotFood.FalafelChickpea,
            StringConstants.FalafelBuckwheat => HotFood.FalafelBuckwheat,
            StringConstants.KebabChicken => HotFood.KebabChicken,
            StringConstants.KebabPork => HotFood.KebabPork,
            _ => null,
        };
    }

    private Lunch? GetLunch(string lunch)
    {
        return lunch switch
        {
            StringConstants.Manager1 => Lunch.Manager1,
            StringConstants.Manager2 => Lunch.Manager2,
            StringConstants.BusinessLady1 => Lunch.BusinessLady1,
            StringConstants.BusinessLady2 => Lunch.BusinessLady2,
            StringConstants.Freelancer1 => Lunch.Freelancer1,
            StringConstants.Freelancer2 => Lunch.Freelancer2,
            StringConstants.Gamer1 => Lunch.Gamer1,
            StringConstants.Gamer2 => Lunch.Gamer2,
            StringConstants.Vegan1 => Lunch.Vegan1,
            StringConstants.Vegan2 => Lunch.Vegan2,
            StringConstants.Vegan3 => Lunch.Vegan3,
            StringConstants.Prince1 => Lunch.Prince1,
            StringConstants.Prince2 => Lunch.Prince2,
            _ => null
        };
    }

    private OrderTime GetOrderTime(string date)
    {
        return date switch
        {
            StringConstants.MorningOrder => OrderTime.Morning,
            StringConstants.DayOrder => OrderTime.Day,
            StringConstants.NightOrder => OrderTime.Night,
            _ => OrderTime.Default
        };
    }
}

public class Order
{
    public readonly OrderTime OrderTime;
    public readonly Lunch? Lunch;
    public readonly HotFood? HotFood;
    public readonly Soup? Soup;
    public readonly Bakery? Bakery;
    public readonly bool WillCoffee;

    public Order(OrderTime orderTime, Lunch? lunch, HotFood? hotFood, Soup? soup, Bakery? bakery, bool willCoffee)
    {
        OrderTime = orderTime;
        Lunch = lunch;
        HotFood = hotFood;
        Soup = soup;
        Bakery = bakery;
        WillCoffee = willCoffee;
    }
}