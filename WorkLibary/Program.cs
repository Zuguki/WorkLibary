using System.Reflection;
using Aspose.Words;
using OfficeOpenXml;
using WorkLibary;
using WorkLibary.Builders.Excel;
using WorkLibary.Lunch;

var startRow = 2;
var maxRow = 100;


var fi = new FileInfo(@"A.xlsx");
var users = new List<User>();

var reader = new ExcelReader(new ExcelPackage(fi).Workbook.Worksheets["Sheet"]);
foreach (var userAndOrders in reader.ReadAllLines())
{
    var userAndLocation = userAndOrders[0].Split(":;:");

    var tuesday = userAndOrders[1].Split(":;:");
    var wednesday = userAndOrders[2].Split(":;:");
    var thursday = userAndOrders[3].Split(":;:");
    var friday = userAndOrders[4].Split(":;:");

    var user = new User(userAndLocation[0], userAndOrders[1]);
    user.AddOrder(tuesday[0], tuesday[1], tuesday[2], tuesday[3], tuesday[4], tuesday[5]);
    user.AddOrder(wednesday[0], wednesday[1], wednesday[2], wednesday[3], wednesday[4], wednesday[5]);
    user.AddOrder(thursday[0], thursday[1], thursday[2], thursday[3], thursday[4], thursday[5]);
    user.AddOrder(friday[0], friday[1], friday[2], friday[3], friday[4], friday[5]);

    users.Add(user);
}

// TODO: OBSOLETE
// using (var p = new ExcelPackage(fi))
// {
//     var ws = p.Workbook.Worksheets["Sheet"];
//     var rows = 100;
//     for (var row = 2; row < rows; row++)
//     {
//         var user = new User((string) ws.Cells[row, 3].Value, ws.Cells[row, 32].Value);
//         user.AddOrder(ws.Cells[row, 5].Value, ws.Cells[row, 6].Value, ws.Cells[row, 7].Value, ws.Cells[row, 8].Value,
//             ws.Cells[row, 9].Value, ws.Cells[row, 10].Value);
//         user.AddOrder(ws.Cells[row, 12].Value, ws.Cells[row, 13].Value, ws.Cells[row, 14].Value, 
//             ws.Cells[row, 15].Value, ws.Cells[row, 16].Value, ws.Cells[row, 17].Value);
//         user.AddOrder(ws.Cells[row, 19].Value, ws.Cells[row, 20].Value, ws.Cells[row, 21].Value, 
//             ws.Cells[row, 22].Value, ws.Cells[row, 23].Value, ws.Cells[row, 24].Value);
//         user.AddOrder(ws.Cells[row, 26].Value, ws.Cells[row, 27].Value, ws.Cells[row, 28].Value, 
//             ws.Cells[row, 29].Value, ws.Cells[row, 30].Value, ws.Cells[row, 31].Value);
//
//         users.Add(user);
//     }
//     
//     p.Save();
// }


var builder = new ExcelBuilder();

foreach (var day in new[] {Days.Tuesday, Days.Wednesday})
{
    GenerateForOrder(day, builder);
    GenerateForKitchen(day, builder);
}

builder.Build();

// using (var package = new ExcelPackage())
// {
//     // foreach (var day in new[] {Days.Tuesday, Days.Wednesday, Days.Thursday, Days.Friday})
//     foreach (var day in new[] {Days.Tuesday, Days.Wednesday})
//     {
//         TryReplaceToLunches((int) day);
//
//         var column = 2;
//         var hotFoods = GetHotFoodsDictionary(users.Where(user => user.Orders.Count > (int) day), (int) day);
//         var soups = GetSoupsDictionary(users.Where(user => user.Orders.Count > (int) day), (int) day);
//         var bakery = GetBakeryDictionary(users.Where(user => user.Orders.Count > (int) day), (int) day);
//         var lunches = GetLunches(users.Where(user => user.Orders.Count > (int) day), (int) day);
//         var lunchesOnceCount = GetLunches2(
//             users.Where(user => user.Orders.Count > (int) day && user.Orders[(int) day].Lunch is not null), (int) day);
//
//         // GenerateUserOrder(package, column - 1, day);
//         // GenerateForOrder(day);
//         // GenerateForKitchen(day);
//
//         // var documentKitchen = new Document();
//         // var documentYandex = new Document();
//         // var builderKitchen = new DocumentBuilder(documentKitchen);
//         // var builderYandex = new DocumentBuilder(documentYandex);
//         // GenerateKitchenDocument(builderKitchen, day);
//         // GenerateYandexDocument(builderYandex, day);
//
//         // documentKitchen.Save($"A{day}Kitchen.docx");
//         // documentYandex.Save($"A{day}Yandex.docx");
//     }
//
//     // package.SaveAs(new FileInfo(@"AAA.xlsx"));
// }

void GenerateYandexDocument(DocumentBuilder builder, Days day)
{
    builder.MoveToDocumentStart();
    builder.Font.Size = 14d;
    builder.Font.Bold = true;
    builder.Writeln(day.GetDescription());
    builder.Font.Bold = false;
    builder.Font.Size = 12d;

    GenerateFor(builder, day, Location.Tramvainaya, OrderTime.Day, "Трамвайный,15 – день! (12:30) Яндекс", false, true);
    GenerateFor(builder, day, Location.Tramvainaya, OrderTime.Night, "Трамвайный,15 – вечер! (15:30) Яндекс", false,
        true);
}

void GenerateKitchenDocument(DocumentBuilder builder, Days day)
{
    builder.MoveToDocumentStart();
    builder.Font.Size = 14d;
    builder.Font.Bold = true;
    builder.Writeln(day.GetDescription());
    builder.Font.Bold = false;
    builder.Font.Size = 12d;

    GenerateFor(builder, day, Location.Vosstaniya, OrderTime.Morning, "Восстания,32 – утро!");
    GenerateFor(builder, day, Location.Tramvainaya, OrderTime.Day, "Трамвайный,15 – день! (12:30)", true);
    // GenerateFor(builder, day, Location.Tramvainaya, OrderTime.Day, "Трамвайный,15 – день! (12:30)");
    GenerateFor(builder, day, Location.Gagarina, OrderTime.Day, "Гагарина 28, Д – день! (12:30)");
    GenerateFor(builder, day, Location.Tramvainaya, OrderTime.Night, "Трамвайный,15 – вечер! (15:30)", true);
    // GenerateFor(builder, day, Location.Tramvainaya, OrderTime.Night, "Трамвайный,15 – вечер! (15:30)");
    GenerateFor(builder, day, Location.Gagarina, OrderTime.Night, "Гагарина 28, Д – вечер! (15:30)");
}

void GenerateFor(DocumentBuilder builder, Days day, Location location, OrderTime time, string text,
    bool withoutLunches = false, bool withoutFood = false)
{
    var counter = 1;
    builder.Font.Bold = true;
    builder.Writeln(text);
    builder.Writeln();

    builder.StartTable();
    AddToTable(builder, "№", "ФИО", "Что заказали");
    builder.Font.Bold = false;

    foreach (var user in users.Where(us => us.Location == location))
    {
        var order = user.Orders[(int) day];
        if (order.OrderTime != time || (withoutLunches && order.Lunch is not null) ||
            (withoutFood && order.Lunch is null))
            continue;

        var hotFoodText = order.HotFood is not null ? order.HotFood!.Value.GetDescription() : "";
        var soupText = order.Soup is not null ? order.Soup!.Value.GetDescription() : "";
        var bakeryText = order.Bakery is not null ? order.Bakery!.Value.GetDescription() : "";

        var right = order.Lunch is not null ? order.Lunch.ToString() : $"{hotFoodText}\t{soupText}\t{bakeryText}";
        AddToTable(builder, counter.ToString(), user.Name, right);

        counter++;
    }
}

void AddToTable(DocumentBuilder builder, string left, string center, string right)
{
    builder.InsertCell();
    builder.Write(left);

    builder.InsertCell();
    builder.Write(center);

    builder.InsertCell();
    builder.Write(right);
    builder.EndRow();
}

void GenerateUserOrder(ExcelPackage p, int column, Days day)
{
    var ws = p.Workbook.Worksheets.Add(day + " Заказы");

    ws.Cells[1, 1].Value = "ФИО";
    ws.Cells[1, 2].Value = "Время доставки";
    ws.Cells[1, 3].Value = "Закажите набор";
    ws.Cells[1, 4].Value = "Закажите горячее";
    ws.Cells[1, 5].Value = "Закажите суп";
    ws.Cells[1, 6].Value = "Закажите десерт";
    ws.Cells[1, 7].Value = "Будет ли кофе?";
    ws.Cells[1, 8].Value = "Локация";

    var row = 2;
    foreach (var user in users)
    {
        ws.Cells[row, column++].Value = user.Name;
        ws.Cells[row, column++].Value = user.Orders[(int) day].OrderTime.GetDescription();
        ws.Cells[row, column++].Value = user.Orders[(int) day].Lunch?.ToString();
        ws.Cells[row, column++].Value = user.Orders[(int) day].HotFood?.GetDescription();
        ws.Cells[row, column++].Value = user.Orders[(int) day].Soup?.GetDescription();
        ws.Cells[row, column++].Value = user.Orders[(int) day].Bakery?.GetDescription();
        ws.Cells[row, column++].Value = user.Orders[(int) day].WillCoffee;
        ws.Cells[row, column++].Value = user.Location.GetDescription();

        row++;
        column = 1;
    }
}

void GenerateForKitchen(Days day, ExcelBuilder builder)
{
    var pageBuilder = builder
        .AddPage(day.GetDescription() + " Заготовки")
        .SetTitles();

    var reader = new ExcelReader(pageBuilder.Worksheet);
    var row = 2;

    pageBuilder.AddCell(1, 2, "Утро");
    pageBuilder.AddCell(1, 3, "День");
    pageBuilder.AddCell(1, 4, "Вечер");

    foreach (var product in reader.ReadCellsById(1, 2))
    {
        pageBuilder.AddCell(row, 2,
            (GetFoodCount(product, OrderTime.Morning, day, true) +
             GetFoodCountInLunch(product, OrderTime.Morning, day, true)).ToString());
        
        pageBuilder.AddCell(row, 3,
            (GetFoodCount(product, OrderTime.Day, day, true) +
             GetFoodCountInLunch(product, OrderTime.Day, day, true)).ToString());
        
        pageBuilder.AddCell(row++, 4,
            (GetFoodCount(product, OrderTime.Night, day, true) +
             GetFoodCountInLunch(product, OrderTime.Night, day, true)).ToString());
    }
}

void GenerateForOrder(Days day, ExcelBuilder builder)
{
    var pageBuilder = builder
        .AddPage(day.GetDescription() + " Заказ")
        .SetTitles();
    
    var reader = new ExcelReader(pageBuilder.Worksheet);
    var row = 2;

    pageBuilder.AddCell(1, 2, "Утренние");
    pageBuilder.AddCell(1, 3, "Дневные");
    pageBuilder.AddCell(1, 4, "Вечерние");
    foreach (var product in reader.ReadCellsById(1, 2))
    {
        pageBuilder.AddCell(row, 2, GetFoodCount(product, OrderTime.Morning, day, true).ToString());
        pageBuilder.AddCell(row, 3, GetFoodCount(product, OrderTime.Day, day, true).ToString());
        pageBuilder.AddCell(row++, 4, GetFoodCount(product, OrderTime.Night, day, true).ToString());
    }
}

int GetFoodCountInLunch(string product, OrderTime orderTime, Days day, bool withYandex = true)
{
    var dayIndex = (int) day;
    
    var hotFood = User.GetHotFood(product);
    if (hotFood is not null)
        return users!.Where(user =>
                user.Orders[dayIndex].OrderTime == orderTime && user.Orders[dayIndex].Lunch is not null)
            .Count(user => user.Orders[dayIndex].Lunch?.HotFood.GetDescription() == product);

    var soup = User.GetSoup(product);
    if (soup is not null)
        return users!.Where(user =>
                user.Orders[dayIndex].OrderTime == orderTime && user.Orders[dayIndex].Lunch is not null)
            .Count(user => user.Orders[dayIndex].Lunch?.Soup.GetDescription() == product);

    var bakery = User.GetSoup(product);
    if (bakery is not null)
        return users!.Where(user =>
                user.Orders[dayIndex].OrderTime == orderTime && user.Orders[dayIndex].Lunch is not null)
            .Count(user => user.Orders[dayIndex].Lunch?.Bakery.GetDescription() == product);

    return 0;
}

// TODO: Change with other yandex
int GetFoodCount(string product, OrderTime orderTime, Days day, bool withYandex = true)
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

    var bakery = User.GetSoup(product);
    if (bakery is not null)
        return users
            .Where(us => us.Orders[(int) day].Bakery is not null && us.Orders[(int) day].OrderTime == orderTime)
            .Count(u => u.Orders[(int) day].Bakery!.Value.GetDescription() == product);

    // Lunch
    if (withYandex)
        return users
            .Where(us =>
                (us.Orders[(int) day].Lunch is not null && us.Orders[(int) day].OrderTime == orderTime)
                && us.Location != Location.Tramvainaya)
            .Count(u => u.Orders[(int) day].Lunch!.Name == product);
    else
        return users
            .Where(us =>
                (us.Orders[(int) day].Lunch is not null && us.Orders[(int) day].OrderTime == orderTime))
            .Count(u => u.Orders[(int) day].Lunch!.Name == product);
}