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

    public WordBuilder AddPage()
    {
        documentBuilder.MoveToDocumentEnd();
        return this;
    }

    
}

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