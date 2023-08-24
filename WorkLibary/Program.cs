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

builder.WordBuilder
    .AddDocument()
    .GenerateYandexDocument(builder.Users, Days.Tuesday, new[] {Location.Tramvainaya})
    .GenerateYandexDocument(builder.Users, Days.Wednesday, new[] {Location.Tramvainaya})
    .GenerateYandexDocument(builder.Users, Days.Thursday, new[] {Location.Tramvainaya})
    .GenerateYandexDocument(builder.Users, Days.Friday, new[] {Location.Tramvainaya})
    .Build("Yandex.docx");

builder.WordBuilder
    .AddDocument()
    .GenerateKitchenDocument(builder.Users, Days.Tuesday, new[] {Location.Tramvainaya})
    .GenerateKitchenDocument(builder.Users, Days.Wednesday, new[] {Location.Tramvainaya})
    .GenerateKitchenDocument(builder.Users, Days.Thursday, new[] {Location.Tramvainaya})
    .GenerateKitchenDocument(builder.Users, Days.Friday, new[] {Location.Tramvainaya})
    .Build("Orders.docx");

