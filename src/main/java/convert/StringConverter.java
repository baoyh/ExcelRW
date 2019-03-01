package convert;

import util.ConvertUtil;

public class StringConverter<R> implements Converter<String, R> {
    private final Class<R> targetType;

    public StringConverter(Class<R> targetType) {
        this.targetType = targetType;
    }

    @Override
    public R convert(String value) {
        return ConvertUtil.stringConvert(value, targetType);
    }
}
