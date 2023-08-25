using Aspose.Words;

namespace WorkLibary.Builders.Word;

public class WordTableBuilder
{
    public WordBuilder WordBuilder { get; }
    private DocumentBuilder documentBuilder { get; }

    public WordTableBuilder(WordBuilder wordBuilder, DocumentBuilder documentBuilder)
    {
        WordBuilder = wordBuilder;
        this.documentBuilder = documentBuilder;
    }

    public WordTableBuilder AddToTable(params string[] items)
    {
        foreach (var item in items)
        {
            documentBuilder.InsertCell();
            documentBuilder.Write(item);
        }
        
        documentBuilder.EndRow();
        return this;
    }
}