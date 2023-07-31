namespace WorkLibary.Lunch;

public class Manager1 : Lunch
{
    public override string Name => StringConstants.Manager1;
    public override HotFood HotFood => HotFood.Beef;
    public override Soup Soup => Soup.MushroomSoup;
    public override Bakery Bakery => Bakery.AppleStrudel;
}