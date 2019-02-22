package exception;

public class ExcelRWException extends Exception {
    public ExcelRWException() {
    }

    public ExcelRWException(String message) {
        super(message);
    }

    public ExcelRWException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelRWException(Throwable cause) {
        super(cause);
    }

    public ExcelRWException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
