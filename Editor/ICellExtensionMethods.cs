using NPOI.SS.UserModel;

namespace Kogane
{
    public static class ICellExtensionMethods
    {
        public static string GetValue( this ICell self )
        {
            if ( self == null ) return string.Empty;

            return self.CellType switch
            {
                CellType.Numeric when DateUtil.IsCellDateFormatted( self ) => self.DateCellValue.ToString(),
                CellType.Numeric                                           => self.NumericCellValue.ToString(),
                CellType.String                                            => self.StringCellValue,
                CellType.Formula                                           => self.CellFormula,
                CellType.Blank                                             => string.Empty,
                CellType.Boolean                                           => self.BooleanCellValue.ToString(),
                CellType.Error                                             => self.ErrorCellValue.ToString(),
                _                                                          => string.Empty,
            };
        }
    }
}