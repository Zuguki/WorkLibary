using Aspose.Words;

namespace WorkLibary.Builders.Word;

public class WordBuilder
{
    private DocumentBuilder documentBuilder;
    private Document document;

    public WordBuilder AddDocument(string title = "")
    {
        document = new Document();
        documentBuilder = new DocumentBuilder(document);
        
        documentBuilder.Font.Size = 16;
        documentBuilder.Bold = true;
        documentBuilder.Writeln(title);
        documentBuilder.Font.Size = 12;
        documentBuilder.Bold = false;
        
        return this;
    }

    public WordBuilder Build(string name)
    {
        document.Save(name);
        return this;
    }
    
    public WordTableBuilder CreateTable(string title)
    {
        var tableBuilder = new WordTableBuilder(this, documentBuilder);
        documentBuilder.Font.Bold = true;
        documentBuilder.Writeln(title);

        documentBuilder.StartTable();
        documentBuilder.Font.Bold = false;
        return tableBuilder;
    }
}