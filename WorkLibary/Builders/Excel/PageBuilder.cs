using OfficeOpenXml;

namespace WorkLibary.Builders.Excel;

public class PageBuilder
{
    public ExcelWorksheet Worksheet { get; }

    public PageBuilder(ExcelWorksheet worksheet)
    {
        Worksheet = worksheet;
    }

    public PageBuilder AddCell(int row, int column, string value)
    {
        Worksheet.Cells[row, column].Value = value;
        return this;
    }

    public PageBuilder SetTitles(bool needHodFood = true, bool needSoup = true, bool needBakery = true,
        bool needLunches = true)
    {
        var row = 2;

        // HotFood
        if (needHodFood)
            SetHotFoodsTitle(ref row);

        // Soup
        if (needSoup)
            SetSoupsTitle(ref row);

        // Bakery
        if (needBakery)
            SetBakeriesTitle(ref row);

        // Lunches
        if (needLunches)
            SetLunchesTitle(ref row);

        return this;
    }

    #region SetFoodTitle

    private void SetHotFoodsTitle(ref int row)
    {
        Worksheet.Cells[row++, 1].Value = HotFood.Pork.GetDescription();
        Worksheet.Cells[row++, 1].Value = HotFood.Beef.GetDescription();
        Worksheet.Cells[row++, 1].Value = HotFood.Chicken.GetDescription();
        Worksheet.Cells[row++, 1].Value = HotFood.Shrimp.GetDescription();
        Worksheet.Cells[row++, 1].Value = HotFood.FalafelBeans.GetDescription();
        Worksheet.Cells[row++, 1].Value = HotFood.FalafelChickpea.GetDescription();
        Worksheet.Cells[row++, 1].Value = HotFood.FalafelBuckwheat.GetDescription();
        Worksheet.Cells[row++, 1].Value = HotFood.KebabChicken.GetDescription();
        Worksheet.Cells[row++, 1].Value = HotFood.KebabPork.GetDescription();
        Worksheet.Cells[row++, 1].Value = HotFood.MeetBallsCheese.GetDescription();
        Worksheet.Cells[row++, 1].Value = HotFood.MeetBallsMushroom.GetDescription();
    }

    private void SetSoupsTitle(ref int row)
    {
        Worksheet.Cells[row++, 1].Value = Soup.SpinachSoup.GetDescription();
        Worksheet.Cells[row++, 1].Value = Soup.MushroomSoup.GetDescription();
        Worksheet.Cells[row++, 1].Value = Soup.PumpkinSoup.GetDescription();
    }

    private void SetBakeriesTitle(ref int row)
    {
        Worksheet.Cells[row++, 1].Value = Bakery.AppleStrudel.GetDescription();
        Worksheet.Cells[row++, 1].Value = Bakery.CarrotCake.GetDescription();
        Worksheet.Cells[row++, 1].Value = Bakery.ChocolateCroissant.GetDescription();
        Worksheet.Cells[row++, 1].Value = Bakery.CottageCheesePie.GetDescription();
        Worksheet.Cells[row++, 1].Value = Bakery.CottageCheeseAndCherryPie.GetDescription();
        Worksheet.Cells[row++, 1].Value = Bakery.RoseWithApplesAndCherries.GetDescription();
    }

    private void SetLunchesTitle(ref int row)
    {
        Worksheet.Cells[row++, 1].Value = StringConstants.Manager1;
        Worksheet.Cells[row++, 1].Value = StringConstants.Manager2;
        Worksheet.Cells[row++, 1].Value = StringConstants.BusinessLady1;
        Worksheet.Cells[row++, 1].Value = StringConstants.BusinessLady2;

        Worksheet.Cells[row++, 1].Value = StringConstants.Freelancer1;
        Worksheet.Cells[row++, 1].Value = StringConstants.Freelancer2;

        Worksheet.Cells[row++, 1].Value = StringConstants.Gamer1;
        Worksheet.Cells[row++, 1].Value = StringConstants.Gamer2;

        Worksheet.Cells[row++, 1].Value = StringConstants.Vegan1;
        Worksheet.Cells[row++, 1].Value = StringConstants.Vegan2;
        Worksheet.Cells[row++, 1].Value = StringConstants.Vegan3;

        Worksheet.Cells[row++, 1].Value = StringConstants.Prince1;
        Worksheet.Cells[row++, 1].Value = StringConstants.Prince2;
    }

    #endregion
}