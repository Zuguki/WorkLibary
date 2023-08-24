using Aspose.Words;

namespace WorkLibary.Builders.Word;

public class WordBuilder
{
    private DocumentBuilder documentBulder;
    private Document document;

    public WordBuilder AddDocument()
    {
        document = new Document();
        documentBulder = new DocumentBuilder(document);
        return this;
    }
    
    public WordBuilder GenerateYandexDocument(Days day, Location[]? locations)
    {
        documentBulder.MoveToDocumentStart();
        documentBulder.Font.Size = 14d;
        documentBulder.Font.Bold = true;
        documentBulder.Writeln(day.GetDescription());
        documentBulder.Font.Bold = false;
        documentBulder.Font.Size = 12d;

        GenerateFor(builder, day, Location.Tramvainaya, OrderTime.Day, "Трамвайный,15 – день! (12:30) Яндекс", false, true);
        GenerateFor(builder, day, Location.Tramvainaya, OrderTime.Night, "Трамвайный,15 – вечер! (15:30) Яндекс", false,
            true);
    }

    void GenerateKitchenDocument(Days day, Location[]? locations)
    {
        documentBulder.MoveToDocumentStart();
        documentBulder.Font.Size = 14d;
        documentBulder.Font.Bold = true;
        documentBulder.Writeln(day.GetDescription());
        documentBulder.Font.Bold = false;
        documentBulder.Font.Size = 12d;

        GenerateFor(builder, day, Location.Vosstaniya, OrderTime.Morning, "Восстания,32 – утро!");
        GenerateFor(builder, day, Location.Tramvainaya, OrderTime.Day, "Трамвайный,15 – день! (12:30)", true);
        // GenerateFor(builder, day, Location.Tramvainaya, OrderTime.Day, "Трамвайный,15 – день! (12:30)");
        GenerateFor(builder, day, Location.Gagarina, OrderTime.Day, "Гагарина 28, Д – день! (12:30)");
        GenerateFor(builder, day, Location.Tramvainaya, OrderTime.Night, "Трамвайный,15 – вечер! (15:30)", true);
        // GenerateFor(builder, day, Location.Tramvainaya, OrderTime.Night, "Трамвайный,15 – вечер! (15:30)");
        GenerateFor(builder, day, Location.Gagarina, OrderTime.Night, "Гагарина 28, Д – вечер! (15:30)");
    }

    public WordBuilder Build(string name)
    {
        document.Save(name);
        return this;
    }
    
    private void GenerateFor(IEnumerable<User> users, Days day, string title)
    {
        var counter = 1;
        documentBulder.Font.Bold = true;
        documentBulder.Writeln(title);
        documentBulder.Writeln();

        documentBulder.StartTable();
        AddToTable(documentBulder, "№", "ФИО", "Что заказали");
        documentBulder.Font.Bold = false;

        foreach (var user in users)
        {
            var order = user.Orders[(int) day];

            var hotFoodText = order.HotFood is not null ? order.HotFood!.Value.GetDescription() : "";
            var soupText = order.Soup is not null ? order.Soup!.Value.GetDescription() : "";
            var bakeryText = order.Bakery is not null ? order.Bakery!.Value.GetDescription() : "";

            var right = order.Lunch is not null ? order.Lunch.ToString() : $"{hotFoodText}\t{soupText}\t{bakeryText}";
            AddToTable(builder, counter.ToString(), user.Name, right);

            counter++;
        }
    }

    private IEnumerable<User> GetUsersWithoutLunchByDay(IEnumerable<User> users, Days day, Location location) =>
        users.Where(user => user.Orders[(int) day].Lunch is null && user.Location == location);

    private IEnumerable<User> GetUsersWithLunchByDay(IEnumerable<User> users, Days day, Location location) =>
        users.Where(user => user.Orders[(int) day].Lunch is not null && user.Location == location);
}