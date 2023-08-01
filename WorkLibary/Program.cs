using System.Reflection;
using OfficeOpenXml;
using WorkLibary;
using WorkLibary.Lunch;

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
        var column = 2;
        var hotFoods = GetHotFoodsDictionary(users.Where(user => user.Orders.Length > (int) day), (int) day);
        var soups = GetSoupsDictionary(users.Where(user => user.Orders.Length > (int) day), (int) day);
        var bakery = GetBakeryDictionary(users.Where(user => user.Orders.Length > (int) day), (int) day);
        var lunches = GetLunches(users.Where(user => user.Orders.Length > (int) day), (int) day);
        var lunchesOnceCount = GetLunches2(
            users.Where(user => user.Orders.Length > (int) day && user.Orders[(int) day].Lunch is not null), (int) day);
        TryReplaceToLunches((int) day);

        GenerateForOrder(p, hotFoods, soups, bakery, lunches, lunchesOnceCount, column, day);
        GenerateForKitchen(p, hotFoods, soups, bakery, lunchesOnceCount, column, day);
    }
    
    p.SaveAs(new FileInfo(@"AAA.xlsx"));
}

void GenerateForKitchen(ExcelPackage p, Dictionary<HotFood, Dictionary<OrderTime, int>> hotFoods,
    Dictionary<Soup, Dictionary<OrderTime, int>> soups, Dictionary<Bakery, Dictionary<OrderTime, int>> bakery,
    Dictionary<string, Dictionary<OrderTime, int>> lunchesOnceCount,
    int column, Days day)
{
        var ws = p.Workbook.Worksheets.Add(day + " Kitchen");
        ws.Cells[1, 1].Value = day.GetDescription();
        ws.Cells[1, 1].Style.Font.Bold = true;
        SetNames(ws);
        
        ws.Cells[1, column].Value = "Утро для кухни";
        SetSumToColumn(ws, column++, OrderTime.Morning, hotFoods, soups, bakery, lunchesOnceCount);
        
        ws.Cells[1, column].Value = "День для кухни";
        SetSumToColumn(ws, column++, OrderTime.Day, hotFoods, soups, bakery, lunchesOnceCount);
        
        ws.Cells[1, column].Value = "Вечер для кухни";
        SetSumToColumn(ws, column++, OrderTime.Night, hotFoods, soups, bakery, lunchesOnceCount);
        
        ws.Cells[1, column].Value = "Итого для кухни";
        SetSumToColumn(ws,  column++, null, hotFoods, soups, bakery, lunchesOnceCount);
}

void GenerateForOrder(ExcelPackage p, Dictionary<HotFood, Dictionary<OrderTime, int>> hotFoods,
    Dictionary<Soup, Dictionary<OrderTime, int>> soups, Dictionary<Bakery, Dictionary<OrderTime, int>> bakery,
    Dictionary<string, Dictionary<OrderTime, int>> lunches,
    Dictionary<string, Dictionary<OrderTime, int>> lunchesOnceCount,
    int column, Days day)
{
        var ws = p.Workbook.Worksheets.Add(day.ToString());
        SetNames(ws);
        
        ws.Cells[1, column].Value = "Утренние заказы";
        SetFoodCount(ws, column++, OrderTime.Morning, hotFoods, soups, bakery, lunches);
        
        ws.Cells[1, column].Value = "Дневные заказы";
        SetFoodCount(ws, column++, OrderTime.Day, hotFoods, soups, bakery, lunches);
        
        ws.Cells[1, column].Value = "Вечерние заказы";
        SetFoodCount(ws, column++, OrderTime.Night, hotFoods, soups, bakery, lunches);

        ws.Cells[1, column].Value = "Утренние заказы без наборов";
        SetLunchOnceCount(ws, column++, OrderTime.Morning, lunchesOnceCount);
        
        ws.Cells[1, column].Value = "Дневные заказы без наборов";
        SetLunchOnceCount(ws, column++, OrderTime.Day, lunchesOnceCount);
        
        ws.Cells[1, column].Value = "Вечерние заказы без наборов";
        SetLunchOnceCount(ws, column++, OrderTime.Night, lunchesOnceCount);

        ws.Cells[1, column].Value = "Утро для кухни";
        SetSumToColumn(ws, column++, OrderTime.Morning, hotFoods, soups, bakery, lunchesOnceCount);
        
        ws.Cells[1, column].Value = "День для кухни";
        SetSumToColumn(ws, column++, OrderTime.Day, hotFoods, soups, bakery, lunchesOnceCount);
        
        ws.Cells[1, column].Value = "Вечер для кухни";
        SetSumToColumn(ws, column++, OrderTime.Night, hotFoods, soups, bakery, lunchesOnceCount);
        
        ws.Cells[1, column].Value = "Итого для кухни";
        SetSumToColumn(ws,  column++, null, hotFoods, soups, bakery, lunchesOnceCount);
}

// TODO: Test
void TryReplaceToLunches(int orderIndex)
{
    foreach (var user in users)
    {
        var ourtype = typeof(Lunch);
        var list = Assembly.GetAssembly(ourtype)?.GetTypes()
                .Where(type => type.IsSubclassOf(ourtype));

        var orders = user.Orders[orderIndex];
        if (orders.Lunch is not null)
            continue;

        foreach (var type in list)
        {
            var lunch = (Lunch) Activator.CreateInstance(type)!;
            if (orders.HotFood == lunch.HotFood
                && orders.Soup == lunch.Soup
                && orders.Bakery == lunch.Bakery)
            {
                orders.HotFood = null;
                orders.Soup = null;
                orders.Bakery = null;
                orders.Lunch = lunch;
            }
        }
    }
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

void SetSumToColumn(ExcelWorksheet ws, int columnTo, OrderTime? orderTime,
    Dictionary<HotFood, Dictionary<OrderTime, int>> hotFoods,
    Dictionary<Soup, Dictionary<OrderTime, int>> soups,
    Dictionary<Bakery, Dictionary<OrderTime, int>> bakeries,
    Dictionary<string, Dictionary<OrderTime, int>> lunches)
{
    for (var row = startRow; row < maxRow; row++)
    {
        var lunchValue = (string) ws.Cells[row, 1].Value;
        if (lunchValue is null)
            return;

        var hotFood = User.GetHotFood(lunchValue);
        var soup = User.GetSoup(lunchValue);
        var bakery = User.GetBakery(lunchValue);

        if (hotFood is not null)
            ws.Cells[row, columnTo].Value = GetDictValue(hotFoods, (HotFood) hotFood, orderTime) +
                                            GetDictValue(lunches, lunchValue, orderTime);
        if (soup is not null)
            ws.Cells[row, columnTo].Value = GetDictValue(soups, (Soup) soup, orderTime) +
                                            GetDictValue(lunches, lunchValue, orderTime);
        if (bakery is not null)
            ws.Cells[row, columnTo].Value = GetDictValue(bakeries, (Bakery) bakery, orderTime) +
                                            GetDictValue(lunches, lunchValue, orderTime);
    }
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
            return;

        ws.Cells[row, columnTo].Value = GetDictValue(lunches, value, orderTime);
    }
}

int GetDictValue<T>(Dictionary<T, Dictionary<OrderTime, int>> dictionary, T tValue, OrderTime? day) 
    where T : notnull
{
    if (dictionary.ContainsKey(tValue) && day is not null && dictionary[tValue].ContainsKey((OrderTime) day))
        return dictionary[tValue][(OrderTime) day];
    if (day is null)
        return GetDictValue(dictionary, tValue, OrderTime.Morning) +
               GetDictValue(dictionary, tValue, OrderTime.Day) +
               GetDictValue(dictionary, tValue, OrderTime.Night);
    return 0;
}

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