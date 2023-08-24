using OfficeOpenXml;
using WorkLibary;
using WorkLibary.Builders.Excel;

var fileInfo = new FileInfo(@"A.xlsx");
var users = new List<User>();

var reader = new ExcelReader(new ExcelPackage(fileInfo).Workbook.Worksheets["Sheet"]);
foreach (var userAndOrders in reader.ReadAllLines())
{
    var userAndLocation = userAndOrders[0].Split(":;:");

    var tuesday = userAndOrders[1].Split(":;:");
    var wednesday = userAndOrders[2].Split(":;:");
    var thursday = userAndOrders[3].Split(":;:");
    var friday = userAndOrders[4].Split(":;:");

    var user = new User(userAndLocation[0], userAndLocation[1]);
    user.AddOrder(tuesday[0], tuesday[1], tuesday[2], tuesday[3], tuesday[4], tuesday[5], Days.Tuesday);
    user.AddOrder(wednesday[0], wednesday[1], wednesday[2], wednesday[3], wednesday[4], wednesday[5], Days.Wednesday);
    user.AddOrder(thursday[0], thursday[1], thursday[2], thursday[3], thursday[4], thursday[5], Days.Thursday);
    user.AddOrder(friday[0], friday[1], friday[2], friday[3], friday[4], friday[5], Days.Friday);
    user.TryReplaceToLunch(new [] {Days.Tuesday, Days.Wednesday, Days.Thursday, Days.Friday});

    users.Add(user);
}

var builder = new ExcelBuilder();

foreach (var day in new[] {Days.Tuesday, Days.Wednesday, Days.Thursday, Days.Friday})
{
    GenerateUserOrder(day, builder);
    GenerateForOrder(day, builder, new [] {Location.Tramvainaya});
    GenerateForKitchen(day, builder, new [] {Location.Tramvainaya});
    
    // GenerateForOrder(day, builder, null);
    // GenerateForKitchen(day, builder, null);
}

builder.Build();

void GenerateForKitchen(Days day, ExcelBuilder excelBuilder, Location[]? yandexLocations)
{
    var pageBuilder = excelBuilder
        .AddPage(day.GetDescription() + " Заготовки")
        .SetTitles(needLunches: false);

    var excelReader = new ExcelReader(pageBuilder.Worksheet);
    var row = 2;

    pageBuilder.AddCell(1, 1, $"{day.GetDescription()}");
    pageBuilder.AddCell(1, 2, "Утро");
    pageBuilder.AddCell(1, 3, "День");
    pageBuilder.AddCell(1, 4, "Вечер");
    pageBuilder.AddCell(1, 5, "Итого");

    foreach (var product in excelReader.ReadCellsById(1, 2))
    {
        var morning = GetFoodCount(product, OrderTime.Morning, day, yandexLocations) +
                      GetFoodCountInLunch(product, OrderTime.Morning, day, yandexLocations);
        
        var dayle = GetFoodCount(product, OrderTime.Day, day, yandexLocations) +
                    GetFoodCountInLunch(product, OrderTime.Day, day, yandexLocations);
        
        var night = GetFoodCount(product, OrderTime.Night, day, yandexLocations) +
                    GetFoodCountInLunch(product, OrderTime.Night, day, yandexLocations);
        
        pageBuilder.AddCell(row, 2, morning.ToString());
        pageBuilder.AddCell(row, 3, dayle.ToString());
        pageBuilder.AddCell(row, 4, night.ToString());
        pageBuilder.AddCell(row++, 5, (morning + dayle + night).ToString());
    }
}

void GenerateForOrder(Days day, ExcelBuilder excelBuilder, Location[]? yandexLocations)
{
    var pageBuilder = excelBuilder
        .AddPage(day.GetDescription() + " Заказ")
        .SetTitles();

    var excelReader = new ExcelReader(pageBuilder.Worksheet);
    var row = 2;

    pageBuilder.AddCell(1, 1, $"{day.GetDescription()}");
    // pageBuilder.AddCell(1, 2, "Утренние");
    pageBuilder.AddCell(1, 2, "Дневные");
    pageBuilder.AddCell(1, 3, "Вечерние");
    foreach (var product in excelReader.ReadCellsById(1, 2))
    {
        var morning = GetFoodCount(product, OrderTime.Morning, day, yandexLocations);
        var dayle = GetFoodCount(product, OrderTime.Day, day, yandexLocations);
        var night = GetFoodCount(product, OrderTime.Night, day, yandexLocations);
        
        pageBuilder.AddCell(row, 2, (morning + dayle).ToString());
        pageBuilder.AddCell(row++, 3, night.ToString());
    }
}

void GenerateUserOrder(Days day, ExcelBuilder excelBuilder)
{
    var column = 1;
    var pageBuilder = excelBuilder
        .AddPage(day.GetDescription() + " Заказы")
        
        .AddCell(1, 1, "ФИО")
        .AddCell(1, 2, "Время доставки")
        .AddCell(1, 3, "Закажите набор")
        .AddCell(1, 4, "Закажите горячее")
        .AddCell(1, 5, "Закажите суп")
        .AddCell(1, 6, "Закажите десерт")
        .AddCell(1, 7, "Будет ли кофе")
        .AddCell(1, 8, "Локация");
        
    var row = 2;
    foreach (var user in users)
    {
        pageBuilder
            .AddCell(row, column++, user.Name ?? "")
            .AddCell(row, column++, user.Orders[(int) day].OrderTime.GetDescription() ?? "")
            .AddCell(row, column++, user.Orders[(int) day].Lunch?.ToString() ?? "")
            .AddCell(row, column++, user.Orders[(int) day].HotFood?.GetDescription() ?? "" )
            .AddCell(row, column++, user.Orders[(int) day].Soup?.GetDescription() ?? "")
            .AddCell(row, column++, user.Orders[(int) day].Bakery?.GetDescription() ?? "")
            .AddCell(row, column++, user.Orders[(int) day].WillCoffee ? "Да" : "Нет")
            .AddCell(row++, column++, user.Location.GetDescription() ?? "");

        column = 1;
    }
}

int GetFoodCountInLunch(string product, OrderTime orderTime, Days day, Location[]? yandexLocations = null)
{
    var dayIndex = (int) day;

    var usersWithoutYandex = users.Where(user =>
        yandexLocations is null || yandexLocations.All(location => location != user.Location));

    var hotFood = User.GetHotFood(product);
    if (hotFood is not null)
        return usersWithoutYandex.Where(user =>
                user.Orders[dayIndex].OrderTime == orderTime && user.Orders[dayIndex].Lunch is not null)
            .Count(user => user.Orders[dayIndex].Lunch?.HotFood.GetDescription() == product);

    var soup = User.GetSoup(product);
    if (soup is not null)
        return usersWithoutYandex.Where(user =>
                user.Orders[dayIndex].OrderTime == orderTime && user.Orders[dayIndex].Lunch is not null)
            .Count(user => user.Orders[dayIndex].Lunch?.Soup.GetDescription() == product);

    var bakery = User.GetBakery(product);
    if (bakery is not null)
        return usersWithoutYandex.Where(user =>
                user.Orders[dayIndex].OrderTime == orderTime && user.Orders[dayIndex].Lunch is not null)
            .Count(user => user.Orders[dayIndex].Lunch?.Bakery.GetDescription() == product);

    return 0;
}

int GetFoodCount(string product, OrderTime orderTime, Days day, Location[]? yandexLocations = null)
{
    var hotFood = User.GetHotFood(product);
    if (hotFood is not null)
    {
        return users
            .Where(us => us.Orders[(int) day].HotFood is not null && us.Orders[(int) day].OrderTime == orderTime)
            .Count(u => u.Orders[(int) day].HotFood!.Value.GetDescription() == product);
    }

    var soup = User.GetSoup(product);
    if (soup is not null)
        return users
            .Where(us => us.Orders[(int) day].Soup is not null && us.Orders[(int) day].OrderTime == orderTime)
            .Count(u => u.Orders[(int) day].Soup!.Value.GetDescription() == product);

    var bakery = User.GetBakery(product);
    if (bakery is not null)
        return users
            .Where(us => us.Orders[(int) day].Bakery is not null && us.Orders[(int) day].OrderTime == orderTime)
            .Count(u => u.Orders[(int) day].Bakery!.Value.GetDescription() == product);

    // Lunch
    return users
        .Where(us =>
            us.Orders[(int) day].Lunch is not null && us.Orders[(int) day].OrderTime == orderTime &&
            (yandexLocations is null || yandexLocations!.All(location => location != us.Location)))
        .Count(u => u.Orders[(int) day].Lunch!.Name == product);
}
