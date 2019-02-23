package read;

import exception.ExcelRWException;
import util.Assert;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;

public abstract class ExcelReader {

    public static class Builder {
        private List<String> sheets;
        private int fromRow = -1;
        private int rowLength = -1;
        private int fromColumn = -1;
        private int columnLength = -1;

        public Builder sheets(List<String> names) throws ExcelRWException {
            Assert.notNull(names);
            this.sheets = names;
            return this;
        }

        public Builder fromRow(int i) throws ExcelRWException {
            Assert.greaterThanZero(i);
            this.fromRow = i;
            return this;
        }

        public Builder rowLength(int i) throws ExcelRWException {
            Assert.greaterThanZero(i);
            this.rowLength = i;
            return this;
        }

        public Builder fromColumn(int i) throws ExcelRWException {
            Assert.greaterThanZero(i);
            this.fromColumn = i;
            return this;
        }

        public Builder columnLength(int i) throws ExcelRWException {
            Assert.greaterThanZero(i);
            this.columnLength = i;
            return this;
        }

        public Builder build() {
            return this;
        }

        public int getRowLength() {
            return rowLength;
        }

        public int getColumnLength() {
            return columnLength;
        }

        public List<String> getSheets() {
            return sheets;
        }

        public int getFromRow() {
            return fromRow;
        }

        public int getFromColumn() {
            return fromColumn;
        }

    }

    public abstract List<List<List<String>>> read(InputStream in) throws ExcelRWException, IOException;
}
