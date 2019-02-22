package util;

import exception.ExcelRWException;

public abstract class Assert {
    public static void greaterThanZero(int i) throws ExcelRWException {
        if(i < 0) {
            throw new ExcelRWException("can not start with zero");
        }
    }

    public static void notNull(Object object) throws ExcelRWException {
        if (object == null) {
            throw new ExcelRWException("this argument is required; it must not be null");
        }
    }
}
