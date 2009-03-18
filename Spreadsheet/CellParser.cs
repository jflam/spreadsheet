using System;
using System.Diagnostics;
namespace Spreadsheet {
    public class Coordinate {
        public string Col;
        public int Row;

        public Coordinate(int row, string column) {
            Row = row;
            Col = column;
        }
    }

    public static class CellParser {
        const int A = (int)'A';
        const int A1 = (int)'A' - 1;

        // A-Z, AA-AZ, BA-BZ ... ZA-ZZ
        public static string GenerateColumnName(int i) {
            if (i < 26) {
                return ((char)(A + i)).ToString();
            } else if (i < 702) {
                char lsd = (char)(A + i % 26);
                char msd = (char)(A1 + i / 26);
                return String.Concat(msd, lsd);
            } else {
                throw new ArgumentException("column index out of bounds: " + i.ToString());
            }
        }

        // [A-Z]\d+ [A-ZA-Z]\d+
        public static Coordinate ParseCellName(string cell) {
            Debug.Assert(cell.Length >= 2);
            Debug.Assert(Char.IsLetter(cell[0]));

            int row;
            string col;
            if (Char.IsLetter(cell[1])) {
                // AA-ZZ case
                row = Int32.Parse(cell.Substring(2)) - 1;
                col = cell.Substring(0, 2);
            } else {
                row = Int32.Parse(cell.Substring(1)) - 1;
                col = cell[0].ToString();
            }

            return new Coordinate(row, col);
        }
    }
}