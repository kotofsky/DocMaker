using DocMaker.Domain;

namespace DocMaker;

public interface IWordBuilder
{
    /// <summary>
    /// Build from the stream
    /// </summary>
    /// <param name="stream">Input stream</param>
    /// <param name="template">Template object</param>
    /// <returns>Result stream</returns>
    Stream Build(Stream stream, DocTemplate template);

    /// <summary>
    /// Build from the file path
    /// </summary>
    /// <param name="templatePath">Input file path</param>
    /// <param name="template">Template object</param>
    /// <returns>byte array</returns>
    byte[] Build(string templatePath, DocTemplate template);

    /// <summary>
    /// Create instance of DocTemplate
    /// </summary>
    /// <returns></returns>
    DocTemplate CreateTemplate();
}