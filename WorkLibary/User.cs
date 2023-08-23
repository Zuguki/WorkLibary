using System.ComponentModel;
using WorkLibary.Lunch;

namespace WorkLibary;

public class User
{
    public readonly string? Name;

    public readonly Location Location;

    public readonly List<Order> Orders = new();

    public User(string? name, object? location)
    {
        Name = name;
        Location = GetLocation((string) location);
    }

    public void AddOrder(object date, object lunch, object hotFood, object soup, object bakery, object willCoffee)
    {
        var time = GetOrderTime((string) date);
        var lunchReady = GetLunch((string) lunch);
        var hotFoodReady = GetHotFood((string) hotFood);
        var soupReady = GetSoup((string) soup);
        var bakeryReady = GetBakery((string) bakery);
        
        Orders.Add(new Order(time, lunchReady, hotFoodReady, soupReady, bakeryReady, ((string) willCoffee) == "Да"));
    }

    public static Bakery? GetBakery(string bakery)
    {
        return bakery switch
        {
            StringConstants.AppleStrudel => Bakery.AppleStrudel,
            StringConstants.CarrotCake => Bakery.CarrotCake,
            StringConstants.ChocolateCroissant => Bakery.ChocolateCroissant,
            StringConstants.CottageCheesePie => Bakery.CottageCheesePie,
            StringConstants.CottageCheeseAndCherryPie => Bakery.CottageCheeseAndCherryPie,
            StringConstants.RoseWithApplesAndCherries => Bakery.RoseWithApplesAndCherries,
            _ => null,
        };
    }

    public static Soup? GetSoup(string soup)
    {
        return soup switch
        {
            StringConstants.SpinachSoup => Soup.SpinachSoup,
            StringConstants.MushroomSoup => Soup.MushroomSoup,
            StringConstants.PumpkinSoup => Soup.PumpkinSoup,
            _ => null,
        };
    }

    public static HotFood? GetHotFood(string hotFood)
    {
        return hotFood switch
        {
            StringConstants.Pork => HotFood.Pork,
            StringConstants.Beef => HotFood.Beef,
            StringConstants.Chicken => HotFood.Chicken,
            StringConstants.Shrimp => HotFood.Shrimp,
            StringConstants.FalafelBeans => HotFood.FalafelBeans,
            StringConstants.FalafelChickpea => HotFood.FalafelChickpea,
            StringConstants.FalafelBuckwheat => HotFood.FalafelBuckwheat,
            StringConstants.KebabChicken => HotFood.KebabChicken,
            StringConstants.KebabPork => HotFood.KebabPork,
            StringConstants.MeetBallsCheese => HotFood.MeetBallsCheese,
            StringConstants.MeetBallsMushroom => HotFood.MeetBallsMushroom,
            _ => null,
        };
    }

    private Lunch.Lunch? GetLunch(string lunch)
    {
        return lunch switch
        {
            StringConstants.Manager1 => new Manager1(),
            StringConstants.Manager2 => new Manager2(),
            StringConstants.BusinessLady1 => new BusinessLady1(),
            StringConstants.BusinessLady2 => new BusinessLady2(),
            StringConstants.Freelancer1 => new Freelancer1(),
            StringConstants.Freelancer2 => new Freelancer2(),
            StringConstants.Gamer1 => new Gamer1(),
            StringConstants.Gamer2 => new Gamer2(),
            StringConstants.Vegan1 => new Vegan1(),
            StringConstants.Vegan2 => new Vegan2(),
            StringConstants.Vegan3 => new Vegan3(),
            StringConstants.Prince1 => new Prince1(),
            StringConstants.Prince2 => new Prince2(),
            _ => null
        };
    }

    private OrderTime GetOrderTime(string date)
    {
        return date switch
        {
            StringConstants.MorningOrder => OrderTime.Morning,
            StringConstants.DayOrder => OrderTime.Day,
            StringConstants.NightOrder => OrderTime.Night,
            _ => OrderTime.Default
        };
    }

    private Location GetLocation(string location)
    {
        return location switch
        {
            StringConstants.Gagarina => Location.Gagarina,
            StringConstants.Vosstaniya => Location.Vosstaniya,
            _ => Location.Tramvainaya
        };
    }
}

public enum Location
{
    [Description(StringConstants.Gagarina)]
    Gagarina,
    
    [Description(StringConstants.Tramvainaya)]
    Tramvainaya,
    
    [Description(StringConstants.Vosstaniya)]
    Vosstaniya
}