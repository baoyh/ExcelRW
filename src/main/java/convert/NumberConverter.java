package convert;

import util.ConvertUtil;

public class NumberConverter<R> implements Converter<Number, R> {

    private final Class<R> targetType;

    public NumberConverter(Class<R> targetType) {
        this.targetType = targetType;
    }

    @Override
    public R convert(Number number) {
        return ConvertUtil.numberConvert(number, targetType);
    }
}
