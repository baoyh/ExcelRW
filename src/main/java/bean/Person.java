package bean;

import annotation.Column;
import annotation.SimpleExcel;
import convert.MyConverter;

@SimpleExcel
public class Person {

    @Column(index = 0)
    private String name;

    @Column(index = 1)
    private String gender;

    @Column(index = 2)
    private String address;

    @Column(index = 3)
    private String phone;

    @Column(index = 4)
    private boolean chinese;

    @Column(index = 6)
    private int count;

    @Column(index = 7, converter = MyConverter.class)
    private String s;

    private String ss;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getGender() {
        return gender;
    }

    public void setGender(String gender) {
        this.gender = gender;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    public String getPhone() {
        return phone;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }

    public boolean isChinese() {
        return chinese;
    }

    public void setChinese(boolean chinese) {
        this.chinese = chinese;
    }

    public int getCount() {
        return count;
    }

    public void setCount(int count) {
        this.count = count;
    }

    public String getS() {
        return s;
    }

    public void setS(String s) {
        this.s = s;
    }

    public String getSs() {
        return ss;
    }

    public void setSs(String ss) {
        this.ss = ss;
    }

    @Override
    public String toString() {
        return "Person{" +
                "name='" + name + '\'' +
                ", gender='" + gender + '\'' +
                ", address='" + address + '\'' +
                ", phone='" + phone + '\'' +
                ", chinese=" + chinese +
                ", count=" + count +
                ", s='" + s + '\'' +
                ", ss='" + ss + '\'' +
                '}' + "\n";
    }
}
