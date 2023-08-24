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

