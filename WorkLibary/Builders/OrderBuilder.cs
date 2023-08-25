using OfficeOpenXml;
using WorkLibary.Builders.Excel;
using WorkLibary.Builders.Word;

namespace WorkLibary.Builders;

public class OrderBuilder
{
    public ExcelBuilder ExcelBuilder { get; } = new();
    public WordBuilder WordBuilder { get; } = new();
    
    public List<User> Users = new();
    
    public OrderBuilder GenerateYandexDocument(Days day, Location[]? yandexLocations)
    {
        WordBuilder.AddDocument(day.GetDescription());
        GenerateDocument(GetUsersWithLunchByDay, Users, day, yandexLocations);
        WordBuilder.Build(day.GetDescription() + "Yandex.docx");
        return this;
    }

    public OrderBuilder GenerateKitchenDocument(Days day, Location[]? yandexLocations)
    {
        WordBuilder.AddDocument(day.GetDescription());
        GenerateDocument(GetUsersWithoutLunchByDay, Users, day, yandexLocations);
        WordBuilder.Build(day.GetDescription() + ".docx");
        return this;
    }

    public OrderBuilder AddOrdersFromWorksheet(string path, string worksheet)
    {
        var excelPackage = new ExcelPackage(path).Workbook.Worksheets[worksheet];
        foreach (var line in ExcelReader.ReadAllLines(excelPackage))
        {
            var splitLine = line.Split(":;:");

            var user = new User(splitLine[2], splitLine[31]);
            user.AddOrder(splitLine[4], splitLine[5], splitLine[6], splitLine[7],
                splitLine[8], splitLine[9], Days.Tuesday);
            user.AddOrder(splitLine[11], splitLine[12], splitLine[13], splitLine[14],
                splitLine[15], splitLine[16], Days.Wednesday);
            user.AddOrder(splitLine[18], splitLine[19], splitLine[20], splitLine[21],
                splitLine[22], splitLine[23], Days.Thursday);
            user.AddOrder(splitLine[25], splitLine[26], splitLine[27], splitLine[28],
                splitLine[29], splitLine[30], Days.Friday);
            user.TryReplaceToLunch(new[] {Days.Tuesday, Days.Wednesday, Days.Thursday, Days.Friday});

            Users.Add(user);
        }

        return this;
    }

    public OrderBuilder GenerateForKitchen(Days day, Location[]? yandexLocations)
    {
        var pageBuilder = ExcelBuilder
            .AddPage(day.GetDescription() + " Заготовки")
            .SetTitles(needLunches: false);
        var row = 2;

        pageBuilder.AddCell(1, 1, $"{day.GetDescription()}");
        pageBuilder.AddCell(1, 2, "Утро");
        pageBuilder.AddCell(1, 3, "День");
        pageBuilder.AddCell(1, 4, "Вечер");
        pageBuilder.AddCell(1, 5, "Итого");

        foreach (var product in ExcelReader.ReadCellsById(pageBuilder.Worksheet, 1, row))
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

        return this;
    }

    public OrderBuilder GenerateForOrder(Days day, Location[]? yandexLocations)
    {
        var pageBuilder = ExcelBuilder
            .AddPage(day.GetDescription() + " Заказ")
            .SetTitles();
        var row = 2;

        pageBuilder.AddCell(1, 1, $"{day.GetDescription()}");
        pageBuilder.AddCell(1, 2, "Дневные");
        pageBuilder.AddCell(1, 3, "Вечерние");
        foreach (var product in ExcelReader.ReadCellsById(pageBuilder.Worksheet, 1, 2))
        {
            var morning = GetFoodCount(product, OrderTime.Morning, day, yandexLocations);
            var dayle = GetFoodCount(product, OrderTime.Day, day, yandexLocations);
            var night = GetFoodCount(product, OrderTime.Night, day, yandexLocations);

            pageBuilder.AddCell(row, 2, (morning + dayle).ToString());
            pageBuilder.AddCell(row++, 3, night.ToString());
        }

        return this;
    }

    public OrderBuilder GenerateUserOrder(Days day)
    {
        var column = 1;
        var pageBuilder = ExcelBuilder
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
        foreach (var user in Users)
        {
            pageBuilder
                .AddCell(row, column++, user.Name ?? "")
                .AddCell(row, column++, user.Orders[(int) day].OrderTime.GetDescription() ?? "")
                .AddCell(row, column++, user.Orders[(int) day].Lunch?.ToString() ?? "")
                .AddCell(row, column++, user.Orders[(int) day].HotFood?.GetDescription() ?? "")
                .AddCell(row, column++, user.Orders[(int) day].Soup?.GetDescription() ?? "")
                .AddCell(row, column++, user.Orders[(int) day].Bakery?.GetDescription() ?? "")
                .AddCell(row, column++, user.Orders[(int) day].WillCoffee ? "Да" : "Нет")
                .AddCell(row++, column++, user.Location.GetDescription() ?? "");

            column = 1;
        }

        return this;
    }

    public OrderBuilder Build(string fileName = "AAA.xlsx")
    {
        ExcelBuilder.Build(fileName);
        return this;
    }
    
    private void GenerateDocument(Func<IEnumerable<User>, Days, Location, Location[], OrderTime, IEnumerable<User>> func,
        IEnumerable<User> users, Days day, Location[]? locations)
    {
        foreach (var location in new [] {Location.Vosstaniya, Location.Tramvainaya, Location.Gagarina})
        {
            GenerateTable(func(users, day, location, locations, OrderTime.Morning), day, location.GetDescription() + "- Утро");
            GenerateTable(func(users, day, location, locations, OrderTime.Day), day, location.GetDescription() + "- День");
            GenerateTable(func(users, day, location, locations, OrderTime.Night), day, location.GetDescription() + "- Вечер");
        }
    }

    private IEnumerable<Location> GetOtherLocations(Location[] currentArray)
    {
        var allLocations = new[] {Location.Tramvainaya, Location.Gagarina, Location.Vosstaniya};
        return allLocations.Where(location => currentArray.All(arr => arr != location));
    }
    
    private void GenerateTable(IEnumerable<User> users, Days day, string? title = null)
    {
        var usersArray = users as User[] ?? users.ToArray();
        if (!usersArray.Any())
            return;
                
        var userId = 1;
        var tableBuilder = WordBuilder.CreateTable(title ?? "");
        tableBuilder.AddToTable("№", "ФИО", "Что заказали");

        foreach (var user in usersArray)
        {
            var order = user.Orders[(int) day];

            var hotFoodText = order.HotFood is not null ? order.HotFood!.Value.GetDescription() : "";
            var soupText = order.Soup is not null ? order.Soup!.Value.GetDescription() : "";
            var bakeryText = order.Bakery is not null ? order.Bakery!.Value.GetDescription() : "";

            var result = order.Lunch is not null 
                ? order.Lunch.ToString() 
                : $"{hotFoodText}\t{soupText}\t{bakeryText}";
            
            tableBuilder.AddToTable(userId.ToString(), user.Name!, result!);
            userId++;
        }
    }

    private IEnumerable<User> GetUsersWithoutLunchByDay(IEnumerable<User> users, Days day, Location location,
        Location[] locations, OrderTime time)
    {
        var orderIndex = (int) day;
        return users.Where(user => user.Location == location && user.Orders[orderIndex].OrderTime == time)
            .Where(user => (user.Orders[orderIndex].Lunch is null && locations.Any(loc => loc == user.Location)) ||
                           locations.All(loc => loc != user.Location));
    }

    private IEnumerable<User> GetUsersWithLunchByDay(IEnumerable<User> users, Days day, Location location,
        Location[] locations, OrderTime time)
    {
        var orderIndex = (int) day;
        return users.Where(user =>
            user.Orders[orderIndex].Lunch is not null && locations.Any(loc => user.Location == loc) &&
            user.Orders[orderIndex].OrderTime == time && user.Location == location);
    }

    private int GetFoodCountInLunch(string product, OrderTime orderTime, Days day, Location[]? yandexLocations = null)
    {
        var dayIndex = (int) day;

        var usersWithoutYandex = Users.Where(user =>
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

    private int GetFoodCount(string product, OrderTime orderTime, Days day, Location[]? yandexLocations = null)
    {
        var hotFood = User.GetHotFood(product);
        if (hotFood is not null)
        {
            return Users
                .Where(us => us.Orders[(int) day].HotFood is not null && us.Orders[(int) day].OrderTime == orderTime)
                .Count(u => u.Orders[(int) day].HotFood!.Value.GetDescription() == product);
        }

        var soup = User.GetSoup(product);
        if (soup is not null)
            return Users
                .Where(us => us.Orders[(int) day].Soup is not null && us.Orders[(int) day].OrderTime == orderTime)
                .Count(u => u.Orders[(int) day].Soup!.Value.GetDescription() == product);

        var bakery = User.GetBakery(product);
        if (bakery is not null)
            return Users
                .Where(us => us.Orders[(int) day].Bakery is not null && us.Orders[(int) day].OrderTime == orderTime)
                .Count(u => u.Orders[(int) day].Bakery!.Value.GetDescription() == product);

        // Lunch
        return Users
            .Where(us =>
                us.Orders[(int) day].Lunch is not null && us.Orders[(int) day].OrderTime == orderTime &&
                (yandexLocations is null || yandexLocations!.All(location => location != us.Location)))
            .Count(u => u.Orders[(int) day].Lunch!.Name == product);
    }
}