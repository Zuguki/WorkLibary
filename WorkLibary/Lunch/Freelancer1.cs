namespace WorkLibary.Lunch;

public class Freelancer1 : Lunch
{
    public override string Name => StringConstants.Freelancer1;
    public override HotFood HotFood => HotFood.Chicken;
    public override Soup Soup => Soup.MushroomSoup;
    public override Bakery Bakery => Bakery.RoseWithApplesAndCherries;
}