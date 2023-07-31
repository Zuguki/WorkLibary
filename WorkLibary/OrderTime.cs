using System.ComponentModel;

namespace WorkLibary;

public enum OrderTime
{
    [Description("Утренний (9:00)")]
    Morning,
    
    [Description("Дневная (12:30)")]
    Day,
    
    [Description("Вечерняя (15:30)")]
    Night,
    Default
}