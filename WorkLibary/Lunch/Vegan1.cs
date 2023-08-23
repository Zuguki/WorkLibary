namespace WorkLibary.Lunch;

public class Vegan1 : Lunch
{
    public override string? Name => StringConstants.Vegan1;
    public override HotFood HotFood => HotFood.FalafelChickpea;
    public override Soup Soup => Soup.SpinachSoup;
    public override Bakery Bakery => Bakery.CottageCheeseAndCherryPie;
}