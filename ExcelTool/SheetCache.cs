using libxl;
using System;

namespace ExcelTool
{
    public class SheetCache
    {
        private int _lastRow = 0, _lastCol = 0;
        private CellType[][] _cellType;
        private string[][] _stringValue;
        private double[][] _numValue;

        public void SaveCachFile(byte[] sourceMd5, string filename)
        {
            SheetCacheMgr.CacheWrite writer = new SheetCacheMgr.CacheWrite();

            writer.AppendInt(lastRow());
            writer.AppendInt(lastCol());
            for (int row = 0; row < lastRow(); ++row)
            {
                for (int col = 0; col < lastCol(); ++col)
                {
                    writer.AppendInt((int)_cellType[row][col]);
                    writer.AppendDouble(_numValue[row][col]);
                    writer.AppendString(_stringValue[row][col]);
                }
            }

            writer.SaveToFile(sourceMd5, filename);
        }

        public SheetCache(byte[] bytes, int head)
        {
            int position = head;

            _lastRow = BitConverter.ToInt32(bytes, position); position += 4;
            _lastCol = BitConverter.ToInt32(bytes, position); position += 4;

            _cellType = new CellType[_lastRow][];
            _stringValue = new string[_lastRow][];
            _numValue = new double[_lastRow][];

            for (int r = 0; r < _lastRow; ++r)
            {
                CellType[] cellTypeRow = new CellType[_lastCol];
                _cellType[r] = cellTypeRow;

                string[] stringValueRow = new string[_lastCol];
                _stringValue[r] = stringValueRow;

                double[] numValueRow = new double[_lastCol];
                _numValue[r] = numValueRow;

                for (int c = 0; c < _lastCol; ++c)
                {
                    cellTypeRow[c] = (CellType)BitConverter.ToInt32(bytes, position); position += 4;
                    numValueRow[c] = BitConverter.ToDouble(bytes, position); position += 4;

                    int strLen = BitConverter.ToInt32(bytes, position); position += 4;
                    if (strLen > 0)
                    {
                        stringValueRow[c] = System.Text.Encoding.Default.GetString(bytes, position, strLen);
                    }
                    position += strLen;
                }
            }
        }

        public SheetCache(Sheet xlSheet)
        {
            _lastRow = xlSheet.lastRow();
            _lastCol = xlSheet.lastCol();

            _cellType = new CellType[_lastRow][];

            _stringValue = new string[_lastRow][];
            _numValue = new double[_lastRow][];

            for (int i = 0; i < _lastRow; ++i)
            {
                CellType[] cellTypeRow = new CellType[_lastCol];
                _cellType[i] = cellTypeRow;

                string[] stringValueRow = new string[_lastCol];
                _stringValue[i] = stringValueRow;

                double[] numValueRow = new double[_lastCol];
                _numValue[i] = numValueRow;

                for (int j = 0; j < _lastCol; ++j)
                {
                    cellTypeRow[j]      = xlSheet.cellType(i, j);

                    stringValueRow[j]   = xlSheet.readStr(i, j);
                    numValueRow[j]      = xlSheet.readNum(i, j);
                }
            }
        }

        public CellType cellType(int row, int col)
        {
            return _cellType[row][col];
        }

        public string readStr(int row, int col)
        {
            return _stringValue[row][col];
        }

        public double readNum(int row, int col)
        {
            return _numValue[row][col];
        }

        public int lastRow()
        {
            return _lastRow;
        }
        
        public int lastCol()
        {
            return _lastCol;
        }
    }
}
