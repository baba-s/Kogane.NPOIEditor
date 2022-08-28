# Kogane NPOI Editor

Unity エディタで NPOI を使えるようにするパッケージ

## 使用例

```csharp
using System.IO;
using Kogane;
using NPOI.XSSF.UserModel;
using UnityEditor;

public class Example
{
    [MenuItem( "Tools/Hoge" )]
    public static void Hoge()
    {
        using var fileStream = new FileStream( "", FileMode.Open, FileAccess.Read, FileShare.ReadWrite );

        var workbook = new XSSFWorkbook( fileStream );
        var sheet    = workbook.GetSheetAt( 0 );
        var row      = sheet.GetRow( 0 );
        var cell     = row.GetCell( 0 );
        var value    = cell.GetValue(); // どんなセルの値も string 型で返します
    }
}
```