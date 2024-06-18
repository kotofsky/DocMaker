using DocMaker.Domain;

namespace DocMaker;

public interface IWordBuilder
{
    /// <summary>
    /// Build from the stream
    /// </summary>
    /// <param name="stream">Template stream</param>
    /// <param name="template">Template object</param>
    /// <returns>Result stream</returns>
    Stream Build(Stream stream, DocTemplate template);

    /// <summary>
    /// Build from the file path
    /// </summary>
    /// <param name="templatePath">Template file path</param>
    /// <param name="template">Template object</param>
    /// <returns>result byte array</returns>
    byte[] Build(string templatePath, DocTemplate template);

    /// <summary>
    /// Create instance of DocTemplate
    /// </summary>
    /// <returns></returns>
    DocTemplate CreateTemplate();

    /// <summary>
    /// Async method to build from the stream
    /// </summary>
    /// <param name="stream">Template stream</param>
    /// <param name="template">Template object</param>
    /// <returns>Result stream</returns>
    Task<Stream> BuildAsync(Stream stream, DocTemplate template);

    /// <summary>
    /// async build from the file path
    /// </summary>
    /// <param name="templatePath">Template file path</param>
    /// <param name="template">Template object</param>
    /// <returns>result byte array</returns>
    Task<byte[]> BuildAsync(string templatePath, DocTemplate template);
}