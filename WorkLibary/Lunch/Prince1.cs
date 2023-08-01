namespace WorkLibary.Lunch;

public class Prince1 : Lunch
{
    public override string Name => StringConstants.Prince1;
    public override HotFood HotFood => HotFood.MeetBallsCheese;
    public override Soup Soup => Soup.Default;
    public override Bakery Bakery => Bakery.CottageCheesePie;
}