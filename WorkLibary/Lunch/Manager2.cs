namespace WorkLibary.Lunch;

public class Manager2 : Lunch
{
    public override string? Name => StringConstants.Manager2;
    public override HotFood HotFood => HotFood.Pork;
    public override Soup Soup => Soup.MushroomSoup;
    public override Bakery Bakery => Bakery.AppleStrudel;
}