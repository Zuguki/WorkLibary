namespace WorkLibary.Lunch;

public class Vegan3 : Lunch
{
    public override string Name => StringConstants.Vegan3;
    public override HotFood HotFood => HotFood.FalafelBeans;
    public override Soup Soup => Soup.SpinachSoup;
    public override Bakery Bakery => Bakery.CottageCheeseAndCherryPie;
}