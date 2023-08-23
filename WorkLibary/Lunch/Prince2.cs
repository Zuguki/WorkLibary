namespace WorkLibary.Lunch;

public class Prince2 : Lunch
{
    public override string? Name => StringConstants.Prince2;
    public override HotFood HotFood => HotFood.MeetBallsMushroom;
    public override Soup Soup => Soup.Default;
    public override Bakery Bakery => Bakery.CottageCheesePie;
}