using System;
using NPOI.SS.UserModel;

namespace libxl
{
    public class Sheet
    {
        private ISheet _sheet;

        public Sheet(ISheet sheet)
        {
            _sheet = sheet;
        }

        public ISheet GetNPOISheet()
        {
            return _sheet;
        }

        public CellType cellType(int row, int col)
        {
            IRow r = _sheet.GetRow(row);
            if (r == null)
            {
                return CellType.CELLTYPE_EMPTY;
            }
            ICell cell = r.GetCell(col);
            if (cell == null)
            {
                return CellType.CELLTYPE_EMPTY;
            }

            NPOI.SS.UserModel.CellType ct = cell.CellType;
            if (ct == NPOI.SS.UserModel.CellType.Formula)
            {
                ct = cell.CachedFormulaResultType;
            }

            switch (ct)
            {
                case NPOI.SS.UserModel.CellType.Numeric:
                    return CellType.CELLTYPE_NUMBER;
                case NPOI.SS.UserModel.CellType.String:
                    return CellType.CELLTYPE_STRING;
                case NPOI.SS.UserModel.CellType.Boolean:
                    return CellType.CELLTYPE_BOOLEAN;
                case NPOI.SS.UserModel.CellType.Blank:
                    return CellType.CELLTYPE_BLANK;
                case NPOI.SS.UserModel.CellType.Error:
                    return CellType.CELLTYPE_ERROR;
                default:
                    return CellType.CELLTYPE_EMPTY;
            }
        }

        public string readStr(int row, int col)
        {
            IRow r = _sheet.GetRow(row);
            if (r == null) return null;
            ICell cell = r.GetCell(col);
            if (cell == null) return null;

            try
            {
                NPOI.SS.UserModel.CellType ct = cell.CellType;
                if (ct == NPOI.SS.UserModel.CellType.Formula)
                {
                    ct = cell.CachedFormulaResultType;
                }

                switch (ct)
                {
                    case NPOI.SS.UserModel.CellType.String:
                        return cell.StringCellValue;
                    case NPOI.SS.UserModel.CellType.Numeric:
                        return cell.NumericCellValue.ToString();
                    case NPOI.SS.UserModel.CellType.Boolean:
                        return cell.BooleanCellValue.ToString();
                    case NPOI.SS.UserModel.CellType.Blank:
                        return null;
                    default:
                        return null;
                }
            }
            catch
            {
                return null;
            }
        }

        public double readNum(int row, int col)
        {
            IRow r = _sheet.GetRow(row);
            if (r == null) return 0.0;
            ICell cell = r.GetCell(col);
            if (cell == null) return 0.0;

            try
            {
                NPOI.SS.UserModel.CellType ct = cell.CellType;
                if (ct == NPOI.SS.UserModel.CellType.Formula)
                {
                    ct = cell.CachedFormulaResultType;
                }

                switch (ct)
                {
                    case NPOI.SS.UserModel.CellType.Numeric:
                        return cell.NumericCellValue;
                    case NPOI.SS.UserModel.CellType.String:
                        double d;
                        if (double.TryParse(cell.StringCellValue, out d))
                            return d;
                        return 0.0;
                    case NPOI.SS.UserModel.CellType.Boolean:
                        return cell.BooleanCellValue ? 1.0 : 0.0;
                    default:
                        return 0.0;
                }
            }
            catch
            {
                return 0.0;
            }
        }

        public bool writeStr(int row, int col, string value)
        {
            IRow r = _sheet.GetRow(row);
            if (r == null) r = _sheet.CreateRow(row);
            ICell cell = r.GetCell(col);
            if (cell == null) cell = r.CreateCell(col);
            cell.SetCellValue(value);
            return true;
        }

        public bool writeNum(int row, int col, double value)
        {
            IRow r = _sheet.GetRow(row);
            if (r == null) r = _sheet.CreateRow(row);
            ICell cell = r.GetCell(col);
            if (cell == null) cell = r.CreateCell(col);
            cell.SetCellValue(value);
            return true;
        }

        public int lastRow()
        {
            return _sheet.LastRowNum + 1;
        }

        public int lastCol()
        {
            int maxCol = 0;
            for (int i = _sheet.FirstRowNum; i <= _sheet.LastRowNum; i++)
            {
                IRow r = _sheet.GetRow(i);
                if (r != null && r.LastCellNum > maxCol)
                {
                    maxCol = r.LastCellNum;
                }
            }
            return maxCol;
        }

        public int firstRow()
        {
            return _sheet.FirstRowNum;
        }

        public int firstCol()
        {
            int minCol = int.MaxValue;
            for (int i = _sheet.FirstRowNum; i <= _sheet.LastRowNum; i++)
            {
                IRow r = _sheet.GetRow(i);
                if (r != null && r.FirstCellNum >= 0 && r.FirstCellNum < minCol)
                {
                    minCol = r.FirstCellNum;
                }
            }
            return minCol == int.MaxValue ? 0 : minCol;
        }
    }
}
