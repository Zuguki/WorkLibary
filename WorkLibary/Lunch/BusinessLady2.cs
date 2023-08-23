namespace WorkLibary.Lunch;

public class BusinessLady2 : Lunch
{
    public override string? Name => StringConstants.BusinessLady2;
    public override HotFood HotFood => HotFood.Chicken;
    public override Soup Soup => Soup.PumpkinSoup;
    public override Bakery Bakery => Bakery.ChocolateCroissant;
}