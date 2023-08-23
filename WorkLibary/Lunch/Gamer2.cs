namespace WorkLibary.Lunch;

public class Gamer2 : Lunch
{
    public override string? Name => StringConstants.Gamer2;
    public override HotFood HotFood => HotFood.KebabChicken;
    public override Soup Soup => Soup.MushroomSoup;
    public override Bakery Bakery => Bakery.CottageCheeseAndCherryPie;
}