using System.ComponentModel;

namespace WorkLibary;

public enum HotFood
{
    [Description(StringConstants.Pork)]
    Pork, 
    
    [Description(StringConstants.Beef)]
    Beef,
    
    [Description(StringConstants.Chicken)]
    Chicken, 
    
    [Description(StringConstants.Shrimp)]
    Shrimp, 
    
    [Description(StringConstants.FalafelBeans)]
    FalafelBeans, 
    
    [Description(StringConstants.FalafelChickpea)]
    FalafelChickpea, 
    
    [Description(StringConstants.FalafelBuckwheat)]
    FalafelBuckwheat, 
    
    [Description(StringConstants.KebabChicken)]
    KebabChicken, 
    
    [Description(StringConstants.KebabPork)]
    KebabPork, 
}