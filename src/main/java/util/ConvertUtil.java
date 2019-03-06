package util;

import exception.ExcelRWException;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

public abstract class ConvertUtil {
    private static final BigInteger LONG_MIN = BigInteger.valueOf(Long.MIN_VALUE);

    private static final BigInteger LONG_MAX = BigInteger.valueOf(Long.MAX_VALUE);

    @SuppressWarnings("unchecked")
    public static <T> T numberConvert(Number number, Class<T> targetClass)
            throws ExcelRWException {

        Assert.notNull(number, "Number must not be null");
        Assert.notNull(targetClass, "Target class must not be null");

        if (targetClass.isInstance(number)) {
            return (T) number;
        }
        else if (Byte.class == targetClass || byte.class == targetClass) {
            long value = number.longValue();
            if (value < Byte.MIN_VALUE || value > Byte.MAX_VALUE) {
                raiseOverflowException(number, targetClass);
            }
            return (T) Byte.valueOf(number.byteValue());
        }
        else if (Short.class == targetClass || short.class == targetClass) {
            long value = number.longValue();
            if (value < Short.MIN_VALUE || value > Short.MAX_VALUE) {
                raiseOverflowException(number, targetClass);
            }
            return (T) Short.valueOf(number.shortValue());
        }
        else if (Integer.class == targetClass || int.class == targetClass) {
            long value = number.longValue();
            if (value < Integer.MIN_VALUE || value > Integer.MAX_VALUE) {
                raiseOverflowException(number, targetClass);
            }
            return (T) Integer.valueOf(number.intValue());
        }
        else if (Long.class == targetClass || long.class == targetClass) {
            BigInteger bigInt = null;
            if (number instanceof BigInteger) {
                bigInt = (BigInteger) number;
            } else if (number instanceof BigDecimal) {
                bigInt = ((BigDecimal) number).toBigInteger();
            }
            if (bigInt != null && (bigInt.compareTo(LONG_MIN) < 0 || bigInt.compareTo(LONG_MAX) > 0)) {
                raiseOverflowException(number, targetClass);
            }
            return (T) Long.valueOf(number.longValue());
        }
        else if (BigInteger.class == targetClass) {
            if (number instanceof BigDecimal) {
                return (T) ((BigDecimal) number).toBigInteger();
            } else {
                return (T) BigInteger.valueOf(number.longValue());
            }
        }
        else if (Float.class == targetClass || float.class == targetClass) {
            return (T) Float.valueOf(number.floatValue());
        }
        else if (Double.class == targetClass || double.class == targetClass) {
            return (T) Double.valueOf(number.doubleValue());
        }
        else if (BigDecimal.class == targetClass) {
            return (T) new BigDecimal(number.toString());
        }
        else if (String.class == targetClass) {
            if (number.toString().endsWith(".0")) {
                return (T) String.valueOf(number.intValue());
            }
            return (T) String.valueOf(number.doubleValue());
        }
        else {
            throw new ExcelRWException("Could not convert number [" + number + "] of type [" +
                    number.getClass().getName() + "] to unsupported target class [" + targetClass.getName() + "]");
        }
    }

    @SuppressWarnings("unchecked")
    public static <T> T stringConvert(String text, Class<T> targetClass) {
        Assert.notNull(text, "Text must not be null");
        Assert.notNull(targetClass, "Target class must not be null");
        String trimmed = trimAllWhitespace(text);

        if (targetClass.isInstance(text)) {
            return (T) text;
        }
        else if (Byte.class == targetClass || byte.class == targetClass) {
            return (T) (isHexNumber(trimmed) ? Byte.decode(trimmed) : Byte.valueOf(trimmed));
        }
        else if (Short.class == targetClass || short.class == targetClass) {
            return (T) (isHexNumber(trimmed) ? Short.decode(trimmed) : Short.valueOf(trimmed));
        }
        else if (Integer.class == targetClass || int.class == targetClass) {
            return (T) (isHexNumber(trimmed) ? Integer.decode(trimmed) : Integer.valueOf(trimmed));
        }
        else if (Long.class == targetClass || long.class == targetClass) {
            return (T) (isHexNumber(trimmed) ? Long.decode(trimmed) : Long.valueOf(trimmed));
        }
        else if (BigInteger.class == targetClass) {
            return (T) (isHexNumber(trimmed) ? decodeBigInteger(trimmed) : new BigInteger(trimmed));
        }
        else if (Float.class == targetClass || float.class == targetClass) {
            return (T) Float.valueOf(trimmed);
        }
        else if (Double.class == targetClass || double.class == targetClass) {
            return (T) Double.valueOf(trimmed);
        }
        else if (BigDecimal.class == targetClass || Number.class == targetClass) {
            return (T) new BigDecimal(trimmed);
        }
        else if (Date.class == targetClass) {
            return (T) parseDate(trimmed);
        }
        else if (Calendar.class == targetClass) {
            Date date = parseDate(trimmed);
            Calendar calendar = Calendar.getInstance();
            calendar.setTime(date);
            return (T) calendar;
        }
        else {
            throw new IllegalArgumentException(
                    "Cannot convert String [" + text + "] to target class [" + targetClass.getName() + "]");
        }
    }

    public static <T> T booleanConvert(Boolean flag, Class<T> targetClass) {
        Assert.notNull(flag, "Flag must not be null");
        Assert.notNull(targetClass, "Target class must not be null");
        if (Boolean.class == targetClass || boolean.class == targetClass) {
            return (T) flag;
        }
        else {
            throw new IllegalArgumentException(
                    "Cannot convert Boolean [" + flag + "] to target class [" + targetClass.getName() + "]");
        }
    }

    private static Date parseDate(String value) {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        Date parse = null;
        try {
            parse = sdf.parse(value);
        } catch (ParseException e) {
            e.printStackTrace();
        }
        return parse;
    }

    private static boolean isHexNumber(String value) {
        int index = (value.startsWith("-") ? 1 : 0);
        return (value.startsWith("0x", index) || value.startsWith("0X", index) || value.startsWith("#", index));
    }

    private static BigInteger decodeBigInteger(String value) {
        int radix = 10;
        int index = 0;
        boolean negative = false;

        // Handle minus sign, if present.
        if (value.startsWith("-")) {
            negative = true;
            index++;
        }

        // Handle radix specifier, if present.
        if (value.startsWith("0x", index) || value.startsWith("0X", index)) {
            index += 2;
            radix = 16;
        }
        else if (value.startsWith("#", index)) {
            index++;
            radix = 16;
        }
        else if (value.startsWith("0", index) && value.length() > 1 + index) {
            index++;
            radix = 8;
        }

        BigInteger result = new BigInteger(value.substring(index), radix);
        return (negative ? result.negate() : result);
    }

    private static String trimAllWhitespace(String str) {
        if (!(str != null && str.length() > 0)) {
            return str;
        }
        int len = str.length();
        StringBuilder sb = new StringBuilder(str.length());
        for (int i = 0; i < len; i++) {
            char c = str.charAt(i);
            if (!Character.isWhitespace(c)) {
                sb.append(c);
            }
        }
        return sb.toString();
    }

    private static void raiseOverflowException(Number number, Class<?> targetClass) {
        throw new ExcelRWException("Could not convert number [" + number + "] of type [" +
                number.getClass().getName() + "] to target class [" + targetClass.getName() + "]: overflow");
    }
}
