namespace WorkLibary.Lunch;

public abstract class Lunch
{
    public abstract string? Name { get; }
    public abstract HotFood HotFood { get; }
    public abstract Soup Soup { get; }
    public abstract Bakery Bakery { get; }

    public override string? ToString() => Name;
}