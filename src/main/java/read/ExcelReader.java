package read;

import exception.ExcelRWException;
import util.Assert;

public abstract class ExcelReader {

    public static class Builder {
        public int fromSheet;
        public int toSheet;
        public int fromRow;
        public int toRow;
        public int fromColumn;
        public int toColumn;

        public Builder fromSheet(int i) throws ExcelRWException {
            Assert.greaterThanZero(i);
            this.fromSheet = i;
            return this;
        }
    }
}
