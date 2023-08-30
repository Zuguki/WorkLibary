using WorkLibary;
using WorkLibary.Builders;

var yandexLocations = new[] {Location.Tramvainaya};
var dayStart = 29;
var daysInMonth = 31;

var builder = new OrderBuilder()
    .AddOrdersFromWorksheet("A.xlsx", "Sheet")

    .GenerateUserOrder(Days.Tuesday)
    .GenerateForOrder(Days.Tuesday, yandexLocations, (dayStart) % daysInMonth)
    .GenerateForKitchen(Days.Tuesday, yandexLocations, (dayStart) % daysInMonth)

    .GenerateUserOrder(Days.Wednesday)
    .GenerateForOrder(Days.Wednesday, yandexLocations, (dayStart + 1) % daysInMonth)
    .GenerateForKitchen(Days.Wednesday, yandexLocations, (dayStart + 1) % daysInMonth)

    .GenerateUserOrder(Days.Thursday)
    .GenerateForOrder(Days.Thursday, yandexLocations, (dayStart + 2) % daysInMonth)
    .GenerateForKitchen(Days.Thursday, yandexLocations, (dayStart + 2) % daysInMonth)

    .GenerateUserOrder(Days.Friday)
    .GenerateForOrder(Days.Friday, yandexLocations, (dayStart + 3) % daysInMonth)
    .GenerateForKitchen(Days.Friday, yandexLocations, (dayStart + 3) % daysInMonth)
    .Build();

builder
    .GenerateYandexDocument(Days.Tuesday, new[] {Location.Tramvainaya}, (dayStart) % daysInMonth)
    .GenerateKitchenDocument(Days.Tuesday, new[] {Location.Tramvainaya}, (dayStart) % daysInMonth)
    
    .GenerateYandexDocument(Days.Wednesday, new[] {Location.Tramvainaya}, (dayStart + 1) % daysInMonth)
    .GenerateKitchenDocument(Days.Wednesday, new[] {Location.Tramvainaya}, (dayStart + 1) % daysInMonth)
    
    .GenerateYandexDocument(Days.Thursday, new[] {Location.Tramvainaya}, (dayStart + 2) % daysInMonth)
    .GenerateKitchenDocument(Days.Thursday, new[] {Location.Tramvainaya}, (dayStart + 2) % daysInMonth)
    
    .GenerateYandexDocument(Days.Friday, new[] {Location.Tramvainaya}, (dayStart + 3) % daysInMonth)
    .GenerateKitchenDocument(Days.Friday, new[] {Location.Tramvainaya}, (dayStart + 3) % daysInMonth);

