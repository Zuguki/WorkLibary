using System.ComponentModel;

namespace WorkLibary;

public enum Days
{
    [Description("Вторник")]
    Tuesday = 0,
    
    [Description("Среда")]
    Wednesday = 1,
    
    [Description("Четверг")]
    Thursday = 2,
    
    [Description("Пятница")]
    Friday = 3
}