# Word Table Fill

A robust ASP.NET Core web application for dynamically filling Word documents (DOCX) with data from JSON payloads. Perfect for generating documents from templates with table population and scalar field replacement.

## Features

- **Template-Based Document Generation**: Use Word templates with placeholders for dynamic content
- **Flexible Table Population**: Fill tables by structured tags (recommended) or by positional order
- **Scalar Field Replacement**: Replace simple placeholders throughout the document (headers, footers, body)
- **Preserves Formatting**: Maintains original document styling and formatting
- **RESTful API**: Clean HTTP endpoint for document generation
- **Built on OpenXML**: Uses the industry-standard `DocumentFormat.OpenXml` library for DOCX manipulation

## Technology Stack

- **Framework**: ASP.NET Core 8.0
- **Language**: C#
- **Key Dependencies**:
  - `DocumentFormat.OpenXml` v3.3.0 - OOXML document manipulation
  - `Swashbuckle.AspNetCore` v6.6.2 - Swagger/OpenAPI documentation
  - `Microsoft.AspNetCore.OpenApi` - Built-in OpenAPI support

## Quick Start

### Prerequisites

- .NET 8.0 SDK or later
- A Word template file (`.docx`) with placeholders

### Installation & Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/DanVens/Word-Table-Fill.git
   cd Word-Table-Fill
   ```

2. Configure your template path in `appsettings.json`:
   ```json
   {
     "Templates": {
       "Word": "Templates/your-template.docx"
     }
   }
   ```

3. Build and run:
   ```bash
   dotnet build
   dotnet run
   ```

4. Open Swagger UI at `http://localhost:5037/swagger` (or `https://localhost:7013/swagger` for HTTPS)

## API Reference

### Endpoint: `POST /api/document/fill`

**Request Body** (`Content-Type: application/json`):

```json
{
  "scalars": {
    "fld_istaiga": "Institution Name",
    "fld_data": "2025-09-01",
    "fld_klase": "4"
  },
  "tableTag": "studentsTbl",
  "rowsByTag": [
    {
      "fld_nr": "1",
      "fld_vardaspav": "John Doe",
      "fld_asmkodas": "39901010001",
      "fld_kodai": "A,V",
      "fld_pastabos": ""
    }
  ]
}
```

**Response**: Binary DOCX file download (`FilledDocument.docx`)

## Usage Patterns

### 1. Fill by Tag (Recommended)

Use structured field names as dictionary keys:

```json
{
  "tableTag": "studentsTbl",
  "rowsByTag": [
    {
      "fld_nr": "1",
      "fld_vardaspav": "Jane Smith",
      "fld_asmkodas": "40102020002"
    },
    {
      "fld_nr": "2",
      "fld_vardaspav": "Bob Johnson",
      "fld_asmkodas": "39905050005"
    }
  ]
}
```

**Best for**: When you have named fields and want robust, maintainable data mapping.

### 2. Fill by Order

Use positional arrays for values:

```json
{
  "tableTag": "studentsTbl",
  "rowsByOrder": [
    ["1", "Jane Smith", "40102020002"],
    ["2", "Bob Johnson", "39905050005"]
  ]
}
```

**Best for**: When column order is fixed and you want minimal payload size.

### 3. Scalar Field Replacement

Replace placeholder text anywhere in the document:

```json
{
  "scalars": {
    "fld_institution": "Lincoln High School",
    "fld_date": "2025-09-01",
    "fld_class": "4B"
  }
}
```

Use all three together for comprehensive document generation:

```json
{
  "scalars": {
    "fld_institution": "Lincoln High School",
    "fld_date": "2025-09-01"
  },
  "tableTag": "studentsTbl",
  "rowsByTag": [
    {
      "fld_nr": "1",
      "fld_name": "John Doe"
    }
  ]
}
```

## Template Preparation

### Creating Templates

1. **Open Word** and create your document structure
2. **Add Placeholders** using Content Controls:
   - Select text to replace
   - Go to **Developer Tab** → **Design Mode**
   - Add field names that match your JSON keys (e.g., `fld_institution`)

3. **For Tables**:
   - Create a template row with Content Controls for each column
   - Tag the table with a bookmark or use the table identification in code
   - Ensure all field names in the row are unique and match your JSON keys

4. **Save** as `.docx` and store in your `Templates/` folder

## Project Structure

```
Word-Table-Fill/
├── WebApplication1/
│   ├── Controllers/
│   │   └── DocumentController.cs       # API endpoint
│   ├── Command/
│   │   └── PostDocumentCommand.cs      # Core document fill logic
│   ├── Dtos/
│   │   └── FillDocumentDto.cs          # Request DTO
│   ├── DocumentGenerator.cs            # Document creation utilities
│   ├── Program.cs                      # ASP.NET Core setup
│   └── WebApplication1.csproj
├── Sample_BY_ORDER.txt                 # Example payload (positional)
├── Sample_BY_TAGS.txt                  # Example payload (tagged)
└── README.md
```

## Code Overview

### PostDocumentCommand

The heart of the application. Key methods:

- **`Execute(Stream templateStream, FillDocumentDto dto)`** - Main entry point; returns filled document as `MemoryStream`
- **`PopulateTableByTag()`** - Fills tables using field name mapping
- **`PopulateTableByOrder()`** - Fills tables using positional arrays
- **`FillScalarsEverywhere()`** - Replaces scalar placeholders throughout document

### FillDocumentDto

Data transfer object for API requests:

```csharp
public class FillDocumentDto
{
    public string? tableTag { get; set; } = "studentsTbl";
    public List<List<string>>? rowsByOrder { get; set; }
    public List<Dictionary<string,string>>? rowsByTag { get; set; }
    public Dictionary<string,string>? scalars { get; set; }
}
```

## Examples

### cURL Request

```bash
curl -X POST "http://localhost:5037/api/document/fill" \
  -H "Content-Type: application/json" \
  -d @Sample_BY_TAGS.txt \
  --output FilledDocument.docx
```

### PowerShell Request

```powershell
$json = Get-Content Sample_BY_TAGS.txt
Invoke-WebRequest -Uri "http://localhost:5037/api/document/fill" `
  -Method POST `
  -Body $json `
  -ContentType "application/json" `
  -OutFile "FilledDocument.docx"
```

## Error Handling

- **Template Not Found**: Returns 404 with message indicating the template path
- **Invalid JSON**: Returns 400 with validation error details
- **Template Issues**: Check that Content Control field names match JSON keys exactly

## Configuration

Edit `WebApplication1/appsettings.json`:

```json
{
  "Templates": {
    "Word": "Templates/your-template.docx"
  },
  "Logging": {
    "LogLevel": {
      "Default": "Information"
    }
  }
}
```

## Development

### Running Locally

- **HTTP**: `http://localhost:5037`
- **HTTPS**: `https://localhost:7013`
- **Swagger UI**: Available at `/swagger` endpoint

### Build & Test

```bash
# Build
dotnet build

# Run
dotnet run --project WebApplication1/WebApplication1.csproj

# Test with sample data
curl -X POST "http://localhost:5037/api/document/fill" \
  -H "Content-Type: application/json" \
  -d @Sample_BY_TAGS.txt \
  --output test-output.docx
```

## License

This project is open source. See LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit issues and pull requests to improve the project.

## Support

For questions or issues, please open a GitHub issue or contact the maintainer.

---

**Created by**: [DanVens](https://github.com/DanVens)  
**Language**: C#  
**Framework**: ASP.NET Core 8.0