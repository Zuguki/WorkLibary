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
        var soups = GetSoupsDictionary(users.Where(user => user.Orders.Length > (int) day), (int) day);
        var bakery = GetBakeryDictionary(users.Where(user => user.Orders.Length > (int) day), (int) day);
        var lunches = GetLunches(users.Where(user => user.Orders.Length > (int) day), (int) day);
        var lunchesOnceCount = GetLunches2(
            users.Where(user => user.Orders.Length > (int) day && user.Orders[(int) day].Lunch is not null), (int) day);

        SetNames(ws);
        
        ws.Cells[1, 2].Value = "Утренние заказы";
        SetFoodCount(ws, 2, OrderTime.Morning, hotFoods, soups, bakery, lunches);
        
        ws.Cells[1, 3].Value = "Дневные заказы";
        SetFoodCount(ws, 3, OrderTime.Day, hotFoods, soups, bakery, lunches);
        
        ws.Cells[1, 4].Value = "Вечерние заказы";
        SetFoodCount(ws, 4, OrderTime.Night, hotFoods, soups, bakery, lunches);

        ws.Cells[1, 5].Value = "Утренние заказы без наборов";
        SetLunchOnceCount(ws, 5, OrderTime.Morning, lunchesOnceCount);
        
        ws.Cells[1, 6].Value = "Дневные заказы без наборов";
        SetLunchOnceCount(ws, 6, OrderTime.Day, lunchesOnceCount);
        
        ws.Cells[1, 7].Value = "Вечерние заказы без наборов";
        SetLunchOnceCount(ws, 7, OrderTime.Night, lunchesOnceCount);

        // ws.Cells[1, 5].Value = "Утренние наборы";
        // SetLunchesCount(ws, 5, OrderTime.Morning, lunches);
        //
        // ws.Cells[1, 6].Value = "Дневные наборы";
        // SetLunchesCount(ws, 6, OrderTime.Day, lunches);
        //
        // ws.Cells[1, 7].Value = "Вечерние наборы";
        // SetLunchesCount(ws, 7, OrderTime.Night, lunches);
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
    
    // Soup
    ws.Cells[row++, 1].Value = Soup.SpinachSoup.GetDescription();
    ws.Cells[row++, 1].Value = Soup.MushroomSoup.GetDescription();
    ws.Cells[row++, 1].Value = Soup.PumpkinSoup.GetDescription();
    
    // Bakery
    ws.Cells[row++, 1].Value = Bakery.AppleStrudel.GetDescription();
    ws.Cells[row++, 1].Value = Bakery.CarrotCake.GetDescription();
    ws.Cells[row++, 1].Value = Bakery.ChocolateCroissant.GetDescription();
    ws.Cells[row++, 1].Value = Bakery.CottageCheesePie.GetDescription();
    ws.Cells[row++, 1].Value = Bakery.CottageCheeseAndCherryPie.GetDescription();
    ws.Cells[row++, 1].Value = Bakery.RoseWithApplesAndCherries.GetDescription();
    
    // Lunches
    ws.Cells[row++, 1].Value = StringConstants.Manager1;
    ws.Cells[row++, 1].Value = StringConstants.Manager2;
    ws.Cells[row++, 1].Value = StringConstants.BusinessLady1;
    ws.Cells[row++, 1].Value = StringConstants.BusinessLady2;
    
    ws.Cells[row++, 1].Value = StringConstants.Freelancer1;
    ws.Cells[row++, 1].Value = StringConstants.Freelancer2;
    
    ws.Cells[row++, 1].Value = StringConstants.Gamer1;
    ws.Cells[row++, 1].Value = StringConstants.Gamer2;
    
    ws.Cells[row++, 1].Value = StringConstants.Vegan1;
    ws.Cells[row++, 1].Value = StringConstants.Vegan2;
    ws.Cells[row++, 1].Value = StringConstants.Vegan3;
}

void SetFoodCount(ExcelWorksheet ws, int columnTo, OrderTime orderTime,
    Dictionary<HotFood, Dictionary<OrderTime, int>> hotFoods,
    Dictionary<Soup, Dictionary<OrderTime, int>> soups,
    Dictionary<Bakery, Dictionary<OrderTime, int>> bakeries,
    Dictionary<string, Dictionary<OrderTime, int>> lunches)
{
    for (var row = startRow; row < maxRow; row++)
    {
        var hotFood = User.GetHotFood((string) ws.Cells[row, 1].Value);
        if (hotFood is not null)
        {
            ws.Cells[row, columnTo].Value = GetDictValue(hotFoods, (HotFood) hotFood, orderTime);
            continue;
        }

        var soup = User.GetSoup((string) ws.Cells[row, 1].Value);
        if (soup is not null)
        {
            ws.Cells[row, columnTo].Value = GetDictValue(soups, (Soup) soup, orderTime);
            continue;
        }
        
        var bakery = User.GetBakery((string) ws.Cells[row, 1].Value);
        if (bakery is not null)
        {
            ws.Cells[row, columnTo].Value = GetDictValue(bakeries, (Bakery) bakery, orderTime);
            continue;
        }
        
        var lunchTitle = (string) ws.Cells[row, 1].Value;
        if (lunchTitle is not null)
        {
            ws.Cells[row, columnTo].Value = GetDictValue(lunches, lunchTitle, orderTime);
            continue;
        }
    }
}

void SetLunchOnceCount(ExcelWorksheet ws, int columnTo, OrderTime orderTime,
    Dictionary<string, Dictionary<OrderTime, int>> lunches)
{
    for (var row = startRow; row < maxRow; row++)
    {
        var value = (string) ws.Cells[row, 1].Value;
        if (value is null)
            continue;

        ws.Cells[row, columnTo].Value = GetDictValue(lunches, value, orderTime);
    }
}

int GetDictValue<T>(Dictionary<T, Dictionary<OrderTime, int>> dictionary, T tValue, OrderTime day) 
    where T : notnull
{
    if (dictionary.ContainsKey(tValue) && dictionary[tValue].ContainsKey(day))
        return dictionary[tValue][day];
    return 0;
}

// // TODO: CHAAAAANGE
// Dictionary<string, Dictionary<OrderTime, int>> GetLunchesOnceCountDictionary(List<User> userEnumerable,
//     int orderIndex, ExcelWorksheet ws)
// {
//     var result = new Dictionary<string, Dictionary<OrderTime, int>>();
//     for (var row = startRow; row < maxRow; row++)
//     {
//         var findValue = (string) ws.Cells[row, 1].Value;
//         if (findValue is null)
//             return result;
//         
//         result.Add(findValue, new Dictionary<OrderTime, int>());
//         
//         var hotFood = User.GetHotFood(findValue);
//         var soup = User.GetSoup(findValue);
//         var bakery = User.GetBakery(findValue);
//         if (hotFood is not null)
//         {
//             result[findValue].Add(OrderTime.Morning,
//                 userEnumerable.Count(user =>
//                     user.Orders[orderIndex].Lunch!.HotFood == hotFood &&
//                     user.Orders[orderIndex].OrderTime == OrderTime.Morning));
//             
//             result[findValue].Add(OrderTime.Day,
//                 userEnumerable.Count(user =>
//                     user.Orders[orderIndex].Lunch!.HotFood == hotFood &&
//                     user.Orders[orderIndex].OrderTime == OrderTime.Day));
//             
//             result[findValue].Add(OrderTime.Night,
//                 userEnumerable.Count(user =>
//                     user.Orders[orderIndex].Lunch!.HotFood == hotFood &&
//                     user.Orders[orderIndex].OrderTime == OrderTime.Night));
//             Console.WriteLine($"HotFood was added: {hotFood.Value.GetDescription()}");
//             continue;
//         }
//
//         if (soup is not null)
//         {
//             result[findValue].Add(OrderTime.Morning,
//                 userEnumerable.Count(user =>
//                     user.Orders[orderIndex].Lunch!.Soup == soup &&
//                     user.Orders[orderIndex].OrderTime == OrderTime.Morning));
//             
//             result[findValue].Add(OrderTime.Day,
//                 userEnumerable.Count(user =>
//                     user.Orders[orderIndex].Lunch!.Soup == soup &&
//                     user.Orders[orderIndex].OrderTime == OrderTime.Day));
//             
//             result[findValue].Add(OrderTime.Night,
//                 userEnumerable.Count(user =>
//                     user.Orders[orderIndex].Lunch!.Soup == soup &&
//                     user.Orders[orderIndex].OrderTime == OrderTime.Night));
//             continue;
//         }
//
//         if (bakery is not null)
//         {
//             result[findValue].Add(OrderTime.Morning,
//                 userEnumerable.Count(user =>
//                     user.Orders[orderIndex].Lunch!.Bakery == bakery &&
//                     user.Orders[orderIndex].OrderTime == OrderTime.Morning));
//             
//             result[findValue].Add(OrderTime.Day,
//                 userEnumerable.Count(user =>
//                     user.Orders[orderIndex].Lunch!.Bakery == bakery &&
//                     user.Orders[orderIndex].OrderTime == OrderTime.Day));
//             
//             result[findValue].Add(OrderTime.Night,
//                 userEnumerable.Count(user =>
//                     user.Orders[orderIndex].Lunch!.Bakery == bakery &&
//                     user.Orders[orderIndex].OrderTime == OrderTime.Night));
//             
//             continue;
//         }
//     }
//     
//     return result;
// }

Dictionary<string, Dictionary<OrderTime, int>> GetLunches2(IEnumerable<User> userEnumerable, int orderIndex)
{
    var result = new Dictionary<string, Dictionary<OrderTime, int>>();
    foreach (var user in userEnumerable)
    {
        var order = user.Orders[orderIndex];
        var lunch = order.Lunch;
        
        if (!result.ContainsKey(lunch!.HotFood.GetDescription()))
            result.Add(lunch!.HotFood.GetDescription(), new Dictionary<OrderTime, int>());
        if (!result.ContainsKey(lunch!.Soup.GetDescription()))
            result.Add(lunch!.Soup.GetDescription(), new Dictionary<OrderTime, int>());
        if (!result.ContainsKey(lunch!.Bakery.GetDescription()))
            result.Add(lunch!.Bakery.GetDescription(), new Dictionary<OrderTime, int>());
        
        if (!result[lunch.HotFood.GetDescription()].ContainsKey(order.OrderTime))
            result[lunch.HotFood.GetDescription()].Add(order.OrderTime, 0);
        
        if (!result[lunch.Soup.GetDescription()].ContainsKey(order.OrderTime))
            result[lunch.Soup.GetDescription()].Add(order.OrderTime, 0);
        
        if (!result[lunch.Bakery.GetDescription()].ContainsKey(order.OrderTime))
            result[lunch.Bakery.GetDescription()].Add(order.OrderTime, 0);

        result[lunch.HotFood.GetDescription()][order.OrderTime]++;
        result[lunch.Soup.GetDescription()][order.OrderTime]++;
        result[lunch.Bakery.GetDescription()][order.OrderTime]++;
    }

    return result;
}

// TODO; Change
Dictionary<string, Dictionary<OrderTime, int>> GetLunches(IEnumerable<User> userEnumerable, int orderIndex)
{
    var result = new Dictionary<string, Dictionary<OrderTime, int>>();
    foreach (var user in userEnumerable)
    {
        var order = user.Orders[orderIndex];
        var lunch = order.Lunch;
        
        if (lunch is null)
            continue;
        
        if (!result.ContainsKey(lunch.Name))
            result.Add(lunch.Name, new Dictionary<OrderTime, int>());
        if (!result[lunch.Name].ContainsKey(order.OrderTime))
            result[lunch.Name].Add(order.OrderTime, 0);

        result[lunch.Name][order.OrderTime]++;
    }

    return result;
}

// TODO: Change
Dictionary<Bakery, Dictionary<OrderTime, int>> GetBakeryDictionary(IEnumerable<User> userEnumerable, int orderIndex)
{
    var bakeries = new Dictionary<Bakery, Dictionary<OrderTime, int>>();
    foreach (var user in userEnumerable)
    {
        var order = user.Orders[orderIndex];
        var bakery = order.Bakery;

        if (bakery is null) 
            continue;
        
        if (!bakeries.ContainsKey((Bakery) bakery))
            bakeries.Add((Bakery) bakery, new Dictionary<OrderTime, int>());
        if (!bakeries[(Bakery) bakery].ContainsKey(order.OrderTime))
            bakeries[(Bakery) bakery].Add(order.OrderTime, 0);

        bakeries[(Bakery) bakery][order.OrderTime]++;
    }

    return bakeries;
}

// TODO: Change
Dictionary<Soup, Dictionary<OrderTime, int>> GetSoupsDictionary(IEnumerable<User> userEnumerable, int orderIndex)
{
    var soups = new Dictionary<Soup, Dictionary<OrderTime, int>>();
    foreach (var user in userEnumerable)
    {
        var order = user.Orders[orderIndex];
        var soup = order.Soup;

        if (soup is null)
            continue;
        
        if (!soups.ContainsKey((Soup) soup))
            soups.Add((Soup) soup, new Dictionary<OrderTime, int>());
        if (!soups[(Soup) soup].ContainsKey(order.OrderTime))
            soups[(Soup) soup].Add(order.OrderTime, 0);

        soups[(Soup) soup][order.OrderTime]++;
    }

    return soups;
}

// TODO: Change
Dictionary<HotFood, Dictionary<OrderTime, int>> GetHotFoodsDictionary(IEnumerable<User> usersEnumerable, int orderIndex)
{
    var hotFoods = new Dictionary<HotFood, Dictionary<OrderTime, int>>();
    foreach (var user in usersEnumerable)
    {
        var order = user.Orders[orderIndex];
        var food = order.HotFood;

        if (food is null) 
            continue;
        
        if (!hotFoods.ContainsKey((HotFood) food))
            hotFoods.Add((HotFood) food, new Dictionary<OrderTime, int>());
        if (!hotFoods[(HotFood) food].ContainsKey(order.OrderTime))
            hotFoods[(HotFood) food].Add(order.OrderTime, 0);

        hotFoods[(HotFood) food][order.OrderTime]++;
    }

    return hotFoods;
}

Console.WriteLine();

public abstract class Lunch
{
    public abstract string Name { get; }
    public abstract HotFood HotFood { get; }
    public abstract Soup Soup { get; }
    public abstract Bakery Bakery { get; }

    public override string ToString() => Name;
}

public class Manager1 : Lunch
{
    public override string Name => StringConstants.Manager1;
    public override HotFood HotFood => HotFood.Beef;
    public override Soup Soup => Soup.MushroomSoup;
    public override Bakery Bakery => Bakery.AppleStrudel;
}

public class Manager2 : Lunch
{
    public override string Name => StringConstants.Manager2;
    public override HotFood HotFood => HotFood.Pork;
    public override Soup Soup => Soup.MushroomSoup;
    public override Bakery Bakery => Bakery.AppleStrudel;
}

public class BusinessLady1 : Lunch
{
    public override string Name => StringConstants.BusinessLady1;
    public override HotFood HotFood => HotFood.Shrimp;
    public override Soup Soup => Soup.PumpkinSoup;
    public override Bakery Bakery => Bakery.ChocolateCroissant;
}

public class BusinessLady2 : Lunch
{
    public override string Name => StringConstants.BusinessLady2;
    public override HotFood HotFood => HotFood.Chicken;
    public override Soup Soup => Soup.PumpkinSoup;
    public override Bakery Bakery => Bakery.ChocolateCroissant;
}

public class Freelancer1 : Lunch
{
    public override string Name => StringConstants.Freelancer1;
    public override HotFood HotFood => HotFood.Chicken;
    public override Soup Soup => Soup.MushroomSoup;
    public override Bakery Bakery => Bakery.RoseWithApplesAndCherries;
}

public class Freelancer2 : Lunch
{
    public override string Name => StringConstants.Freelancer2;
    public override HotFood HotFood => HotFood.KebabPork;
    public override Soup Soup => Soup.MushroomSoup;
    public override Bakery Bakery => Bakery.RoseWithApplesAndCherries;
}

public class Gamer1 : Lunch
{
    public override string Name => StringConstants.Gamer1;
    public override HotFood HotFood => HotFood.KebabPork;
    public override Soup Soup => Soup.MushroomSoup;
    public override Bakery Bakery => Bakery.CottageCheeseAndCherryPie;
}

public class Gamer2 : Lunch
{
    public override string Name => StringConstants.Gamer2;
    public override HotFood HotFood => HotFood.KebabChicken;
    public override Soup Soup => Soup.MushroomSoup;
    public override Bakery Bakery => Bakery.CottageCheeseAndCherryPie;
}

public class Vegan1 : Lunch
{
    public override string Name => StringConstants.Vegan1;
    public override HotFood HotFood => HotFood.FalafelChickpea;
    public override Soup Soup => Soup.SpinachSoup;
    public override Bakery Bakery => Bakery.CottageCheeseAndCherryPie;
}

public class Vegan2 : Lunch
{
    public override string Name => StringConstants.Vegan2;
    public override HotFood HotFood => HotFood.FalafelBuckwheat;
    public override Soup Soup => Soup.SpinachSoup;
    public override Bakery Bakery => Bakery.CottageCheeseAndCherryPie;
}

public class Vegan3 : Lunch
{
    public override string Name => StringConstants.Vegan3;
    public override HotFood HotFood => HotFood.FalafelBeans;
    public override Soup Soup => Soup.SpinachSoup;
    public override Bakery Bakery => Bakery.CottageCheeseAndCherryPie;
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
            StringConstants.Manager1 => new Manager1(),
            StringConstants.Manager2 => new Manager2(),
            StringConstants.BusinessLady1 => new BusinessLady1(),
            StringConstants.BusinessLady2 => new BusinessLady2(),
            StringConstants.Freelancer1 => new Freelancer1(),
            StringConstants.Freelancer2 => new Freelancer2(),
            StringConstants.Gamer1 => new Gamer1(),
            StringConstants.Gamer2 => new Gamer2(),
            StringConstants.Vegan1 => new Vegan1(),
            StringConstants.Vegan2 => new Vegan2(),
            StringConstants.Vegan3 => new Vegan3(),
            // StringConstants.Prince1 => ,
            // StringConstants.Prince2 => Lunch.Prince2,
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