using DocMaker.Domain;
using DocMaker.Services;
using DocumentFormat.OpenXml.Packaging;

namespace DocMaker;

public class WordBuilder : IWordBuilder
{
    private readonly WordElementsService _wordsService;
    public WordBuilder()
    {
        _wordsService = new WordElementsService();
    }

    /// <inheritdoc />
    public DocTemplate CreateTemplate()
    {
        return new DocTemplate();
    }

    /// <inheritdoc />
    public Stream Build(Stream stream, DocTemplate template)
    {
        return Generate(stream, template);
    }

    /// <inheritdoc />
    public byte[] Build(string templatePath, DocTemplate template)
    {
        if (!File.Exists(templatePath))
            throw new FileNotFoundException($"File not found at this path: {templatePath}");

        return Generate(templatePath, template);
    }

    private Stream Generate(Stream stream, DocTemplate template)
    {
        var resultStream = new MemoryStream();
        stream.CopyTo(resultStream);
        resultStream.Seek(0, SeekOrigin.Begin);

        return WriteIntoFileStream(resultStream, template);
    }

    private byte[] Generate(string templatePath, DocTemplate template)
    {
        byte[] templateData = File.ReadAllBytes(templatePath);

        var templateStream = new MemoryStream();
        templateStream.Write(templateData, 0, templateData.Length);
        templateStream.Seek(0, SeekOrigin.Begin);

        var resultStream = WriteIntoFileStream(templateStream, template);

        return resultStream.ToArray();
    }

    private MemoryStream WriteIntoFileStream(MemoryStream stream, DocTemplate template)
    {
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            foreach (var key in template.FieldsCollection.Keys)
            {
                _wordsService.SetContentControls(doc, key, template.FieldsCollection[key]);
            }

            if (template.Tables.Length > 0)
                _wordsService.ProcessWordTables(doc, template.Tables);
        }

        return stream;
    }

    #region async methods

    /// <inheritdoc />
    public async Task<Stream> BuildAsync(Stream stream, DocTemplate template)
    {
        return await GenerateAsync(stream, template);
    }

    /// <inheritdoc />
    public async Task<byte[]> BuildAsync(string templatePath, DocTemplate template)
    {
        if (!File.Exists(templatePath))
            throw new FileNotFoundException($"File not found at this path: {templatePath}");

        return await GenerateAsync(templatePath, template);
    }

    private async Task<Stream> GenerateAsync(Stream stream, DocTemplate template)
    {
        var resultStream = new MemoryStream();
        await stream.CopyToAsync(resultStream);
        resultStream.Seek(0, SeekOrigin.Begin);

        return WriteIntoFileStream(resultStream, template);
    }

    private async Task<byte[]> GenerateAsync(string templatePath, DocTemplate template)
    {
        byte[] templateData = await File.ReadAllBytesAsync(templatePath);

        var templateStream = new MemoryStream();
        await templateStream.WriteAsync(templateData, 0, templateData.Length);
        templateStream.Seek(0, SeekOrigin.Begin);

        var resultStream = WriteIntoFileStream(templateStream, template);

        return resultStream.ToArray();
    }


    #endregion
}