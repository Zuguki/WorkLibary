namespace WorkLibary.Lunch;

public class Vegan2 : Lunch
{
    public override string Name => StringConstants.Vegan2;
    public override HotFood HotFood => HotFood.FalafelBuckwheat;
    public override Soup Soup => Soup.SpinachSoup;
    public override Bakery Bakery => Bakery.CottageCheeseAndCherryPie;
}