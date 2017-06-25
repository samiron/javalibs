package exp.samiron.regex;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created by samir on 6/25/2017.
 */
public class Simplest {
    public static void main(String[] args) {
        Pattern p = Pattern.compile("([A-Z]+)(\\d+)");
        Matcher m = p.matcher("A1".trim());
        boolean b = m.matches();
        System.out.println(m.group(2));
        System.out.println(m.group(1));


        String col = m.group(1);
        int l = col.length();
        int colnum = 0;
        for(char c : col.toCharArray()){
            l--;
            colnum += ((int)Math.pow(26,l))*(c - 'A' + 1);
        }
        System.out.println("Column Number: " + colnum);
    }
}
