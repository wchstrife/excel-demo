/**
 * Created by wangchenghao on 2017/8/16.
 */
public class test {

    public static void main(String[] args){
        String target = "<col index=\"A\" width='17em'></col>";
        String unit = target.replaceAll("[0-9]", "");
        String value = target.replace(unit, "");
        System.out.println(unit);
        System.out.println(value);
    }
}
