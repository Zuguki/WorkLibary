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

    public WordBuilder GenerateYandexDocument(IEnumerable<User> users, Days day, Location[]? locations) =>
        GenerateDocument(GetUsersWithLunchByDay, users, day, locations);

    public WordBuilder GenerateKitchenDocument(IEnumerable<User> users, Days day, Location[]? locations) =>
        GenerateDocument(GetUsersWithoutLunchByDay, users, day, locations);

    private WordBuilder GenerateDocument(Func<IEnumerable<User>, Days, Location, OrderTime, IEnumerable<User>> func,
        IEnumerable<User> users, Days day, Location[]? locations)
    {
        documentBulder.MoveToDocumentStart();
        documentBulder.Font.Size = 14d;
        documentBulder.Font.Bold = true;
        documentBulder.Writeln(day.GetDescription());
        documentBulder.Font.Bold = false;
        documentBulder.Font.Size = 12d;

        foreach (var location in locations)
        {
            GenerateTable(func(users, day, location, OrderTime.Morning), day, location.GetDescription() + "- Утро");
            GenerateTable(func(users, day, location, OrderTime.Day), day, location.GetDescription() + "- День");
            GenerateTable(func(users, day, location, OrderTime.Night), day, location.GetDescription() + "- Ночь");
        }

        return this;
    }

    public WordBuilder Build(string name)
    {
        document.Save(name);
        return this;
    }
    
    private void GenerateTable(IEnumerable<User> users, Days day, string title)
    {
        if (!users.Any())
            return;
                
        var userId = 1;
        documentBulder.Font.Bold = true;
        documentBulder.Writeln(title);
        documentBulder.Writeln();

        documentBulder.StartTable();
        AddToTable("№", "ФИО", "Что заказали");
        documentBulder.Font.Bold = false;

        foreach (var user in users)
        {
            var order = user.Orders[(int) day];

            var hotFoodText = order.HotFood is not null ? order.HotFood!.Value.GetDescription() : "";
            var soupText = order.Soup is not null ? order.Soup!.Value.GetDescription() : "";
            var bakeryText = order.Bakery is not null ? order.Bakery!.Value.GetDescription() : "";

            var result = order.Lunch is not null 
                ? order.Lunch.ToString() 
                : $"{hotFoodText}\t{soupText}\t{bakeryText}";
            
            AddToTable(userId.ToString(), user.Name!, result!);
            userId++;
        }
    }

    private IEnumerable<User> GetUsersWithoutLunchByDay(IEnumerable<User> users, Days day, Location location, OrderTime time) =>
        users.Where(user =>
            ((user.Orders[(int) day].Lunch is null && user.Location == location) || user.Location != location) &&
            user.Orders[(int) day].OrderTime == time);

    private IEnumerable<User> GetUsersWithLunchByDay(IEnumerable<User> users, Days day, Location location, OrderTime time) =>
        users.Where(user => user.Orders[(int) day].Lunch is not null && user.Location == location &&
                            user.Orders[(int) day].OrderTime == time);
    
    private void AddToTable(string left, string center, string right)
    {
        documentBulder.InsertCell();
        documentBulder.Write(left); 
    
        documentBulder.InsertCell();
        documentBulder.Write(center);
    
        documentBulder.InsertCell();
        documentBulder.Write(right); 
        documentBulder.EndRow();
    } 
}