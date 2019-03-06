package convert;

import util.ConvertUtil;

public class BooleanConverter<R> implements Converter<Boolean, R> {

    private Class<R> target;

    public BooleanConverter(Class<R> target) {
        this.target = target;
    }

    @Override
    public R convert(Boolean aBoolean) {
        return ConvertUtil.booleanConvert(aBoolean, target);
    }
}
