package read;

import exception.ExcelRWException;
import util.Assert;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;

public abstract class ExcelReader {

    public static class Builder {
        private List<String> sheets;
        private int fromRow;
        private int toRow;
        private short fromColumn;
        private short toColumn;

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

        public Builder toRow(int i) {
            this.toRow = i;
            return this;
        }

        public Builder fromColumn(short i) throws ExcelRWException {
            Assert.greaterThanZero(i);
            this.fromColumn = i;
            return this;
        }

        public Builder toColumn(short i) {
            this.toColumn = i;
            return this;
        }

        public List<String> getSheets() {
            return sheets;
        }

        public int getFromRow() {
            return fromRow;
        }

        public int getToRow() {
            return toRow;
        }

        public short getFromColumn() {
            return fromColumn;
        }

        public short getToColumn() {
            return toColumn;
        }
    }

    public abstract List<List<List<String>>> read(InputStream in, Builder builder) throws ExcelRWException, IOException;
}
