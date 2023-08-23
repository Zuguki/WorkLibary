namespace WorkLibary.Lunch;

public class BusinessLady1 : Lunch
{
    public override string? Name => StringConstants.BusinessLady1;
    public override HotFood HotFood => HotFood.Shrimp;
    public override Soup Soup => Soup.PumpkinSoup;
    public override Bakery Bakery => Bakery.ChocolateCroissant;
}