using WorkLibary;
using WorkLibary.Builders;

var yandexLocations = new[] {Location.Tramvainaya};

var builder = new OrderBuilder()
    .AddOrdersFromWorksheet("A.xlsx", "Sheet")

    .GenerateUserOrder(Days.Tuesday)
    .GenerateForOrder(Days.Tuesday, yandexLocations)
    .GenerateForKitchen(Days.Tuesday, yandexLocations)

    .GenerateUserOrder(Days.Wednesday)
    .GenerateForOrder(Days.Wednesday, yandexLocations)
    .GenerateForKitchen(Days.Wednesday, yandexLocations)

    .GenerateUserOrder(Days.Thursday)
    .GenerateForOrder(Days.Thursday, yandexLocations)
    .GenerateForKitchen(Days.Thursday, yandexLocations)

    .GenerateUserOrder(Days.Friday)
    .GenerateForOrder(Days.Friday, yandexLocations)
    .GenerateForKitchen(Days.Friday, yandexLocations)
    .Build();

// var reader = new ExcelReader(new ExcelPackage(fileInfo).Workbook.Worksheets["Sheet"]);

// foreach (var userAndOrders in reader.ReadAllLines())
// {
//     var userAndLocation = userAndOrders[0].Split(":;:");
//
//     var tuesday = userAndOrders[1].Split(":;:");
//     var wednesday = userAndOrders[2].Split(":;:");
//     var thursday = userAndOrders[3].Split(":;:");
//     var friday = userAndOrders[4].Split(":;:");
//
//     var user = new User(userAndLocation[0], userAndLocation[1]);
//     user.AddOrder(tuesday[0], tuesday[1], tuesday[2], tuesday[3], tuesday[4], tuesday[5], Days.Tuesday);
//     user.AddOrder(wednesday[0], wednesday[1], wednesday[2], wednesday[3], wednesday[4], wednesday[5], Days.Wednesday);
//     user.AddOrder(thursday[0], thursday[1], thursday[2], thursday[3], thursday[4], thursday[5], Days.Thursday);
//     user.AddOrder(friday[0], friday[1], friday[2], friday[3], friday[4], friday[5], Days.Friday);
//     user.TryReplaceToLunch(new [] {Days.Tuesday, Days.Wednesday, Days.Thursday, Days.Friday});
//
//     users.Add(user);
// }

// var builder = new ExcelBuilder();

// foreach (var day in new[] {Days.Tuesday, Days.Wednesday, Days.Thursday, Days.Friday})
// {
//     GenerateUserOrder(day, builder);
//     GenerateForOrder(day, builder, new[] {Location.Tramvainaya});
//     GenerateForKitchen(day, builder, new[] {Location.Tramvainaya});
//
//     // GenerateForOrder(day, builder, null);
//     // GenerateForKitchen(day, builder, null);
// }
//
// builder.Build();
