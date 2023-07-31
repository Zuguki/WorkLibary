namespace WorkLibary.Lunch;

public class Gamer1 : Lunch
{
    public override string Name => StringConstants.Gamer1;
    public override HotFood HotFood => HotFood.KebabPork;
    public override Soup Soup => Soup.MushroomSoup;
    public override Bakery Bakery => Bakery.CottageCheeseAndCherryPie;
}