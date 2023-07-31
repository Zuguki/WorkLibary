using System.ComponentModel;
using System.Reflection;
using OfficeOpenXml;
using WorkLibary;

var fi = new FileInfo(@"A.xlsx");
var users = new List<User>();
var startRow = 2;
var maxRow = 100;

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
    // foreach (var day in new[] {Days.Tuesday, Days.Wednesday, Days.Thursday, Days.Friday})
    foreach (var day in new[] {Days.Tuesday})
    {
        var column = 1;
        var row = 1;
        var ws = p.Workbook.Worksheets.Add(day.ToString());
        var hotFoods = GetHotFoodsDictionary(users.Where(user => user.Orders.Length > (int) day), (int) day);
        SetNames(ws);
        
        ws.Cells[1, 2].Value = "Утренние заказы";
        SetHotFoodCount(ws, 2, OrderTime.Morning, hotFoods);
        
        ws.Cells[1, 3].Value = "Дневные заказы";
        SetHotFoodCount(ws, 3, OrderTime.Day, hotFoods);
        
        ws.Cells[1, 4].Value = "Вечерние заказы";
        SetHotFoodCount(ws, 4, OrderTime.Night, hotFoods);
    }
    
    p.SaveAs(new FileInfo(@"AAA.xlsx"));
}

void SetNames(ExcelWorksheet ws)
{
    var row = startRow;
    
    // HotFood
    ws.Cells[row++, 1].Value = HotFood.Pork.GetDescription();
    ws.Cells[row++, 1].Value = HotFood.Beef.GetDescription();
    ws.Cells[row++, 1].Value = HotFood.Chicken.GetDescription();
    ws.Cells[row++, 1].Value = HotFood.Shrimp.GetDescription();
    ws.Cells[row++, 1].Value = HotFood.FalafelBeans.GetDescription();
    ws.Cells[row++, 1].Value = HotFood.FalafelChickpea.GetDescription();
    ws.Cells[row++, 1].Value = HotFood.FalafelBuckwheat.GetDescription();
    ws.Cells[row++, 1].Value = HotFood.KebabChicken.GetDescription();
    ws.Cells[row++, 1].Value = HotFood.KebabPork.GetDescription();
}

void SetHotFoodCount(ExcelWorksheet ws, int columnTo, OrderTime orderTime,
    Dictionary<HotFood, Dictionary<OrderTime, int>> hotFoods)
{
    for (var row = startRow; row < maxRow; row++)
    {
        var value = User.GetHotFood((string) ws.Cells[row, 1].Value);
        if (value is null)
            return;
        
        ws.Cells[row, columnTo].Value = GetFoodValue(hotFoods, (HotFood) User.GetHotFood((string) ws.Cells[row, 1].Value), orderTime);
    }
}

int GetFoodValue(Dictionary<HotFood, Dictionary<OrderTime, int>> hotFoods, HotFood food, OrderTime day)
{
    if (hotFoods.ContainsKey(food) && hotFoods[food].ContainsKey(day))
        return hotFoods[food][day];
    return 0;
}

Dictionary<HotFood, Dictionary<OrderTime, int>> GetHotFoodsDictionary(IEnumerable<User> usersEnumerable, int orderIndex)
{
    var hotFoods = new Dictionary<HotFood, Dictionary<OrderTime, int>>();
    foreach (var user in usersEnumerable)
    {
        var order = user.Orders[orderIndex];
        var food = order.HotFood;
        
        if (food is not null)
        {
            if (!hotFoods.ContainsKey((HotFood) food))
                hotFoods.Add((HotFood) food, new Dictionary<OrderTime, int>());
            if (!hotFoods[(HotFood) food].ContainsKey(order.OrderTime))
                hotFoods[(HotFood) food].Add(order.OrderTime, 0);

            hotFoods[(HotFood) food][order.OrderTime]++;
        }
    }

    return hotFoods;
}

Console.WriteLine();

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

public enum Bakery
{
    [Description(StringConstants.AppleStrudel)]
    AppleStrudel, 
    
    [Description(StringConstants.CarrotCake)]
    CarrotCake, 
    
    [Description(StringConstants.ChocolateCroissant)]
    ChocolateCroissant, 
    
    [Description(StringConstants.CottageCheesePie)]
    CottageCheesePie, 
    
    [Description(StringConstants.CottageCheeseAndCherryPie)]
    CottageCheeseAndCherryPie, 
    
    [Description(StringConstants.RoseWithApplesAndCherries)]
    RoseWithApplesAndCherries, 
}

public class User
{
    public readonly string Name;

    public readonly Order[] Orders = new Order[5];

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
        
        Orders[0] = new Order(time, lunchReady, hotFoodReady, soupReady, bakeryReady, ((string) willCoffee) == "Да");
    }

    public static Bakery? GetBakery(string bakery)
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

    public static Soup? GetSoup(string soup)
    {
        return soup switch
        {
            StringConstants.SpinachSoup => Soup.SpinachSoup,
            StringConstants.MushroomSoup => Soup.MushroomSoup,
            StringConstants.PumpkinSoup => Soup.PumpkinSoup,
            _ => null,
        };
    }

    public static HotFood? GetHotFood(string hotFood)
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

public static class EnumExt
{
    public static string GetDescription<T>(this T enumerationValue)
        where T : struct
    {
        Type type = enumerationValue.GetType();
        if (!type.IsEnum)
        {
            throw new ArgumentException("EnumerationValue must be of Enum type", "enumerationValue");
        }

        //Tries to find a DescriptionAttribute for a potential friendly name
        //for the enum
        MemberInfo[] memberInfo = type.GetMember(enumerationValue.ToString());
        if (memberInfo != null && memberInfo.Length > 0)
        {
            object[] attrs = memberInfo[0].GetCustomAttributes(typeof(DescriptionAttribute), false);

            if (attrs != null && attrs.Length > 0)
            {
                //Pull out the description value
                return ((DescriptionAttribute) attrs[0]).Description;
            }
        }

        //If we have no description attribute, just return the ToString of the enum
        return enumerationValue.ToString();
    }
}
