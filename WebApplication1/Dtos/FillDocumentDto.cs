public class FillDocumentDto
{
    public string? tableTag { get; set; } = "studentsTbl";
    public List<List<string>>? rowsByOrder { get; set; }
    public List<Dictionary<string,string>>? rowsByTag { get; set; }
    public Dictionary<string,string>? scalars { get; set; }
    
}