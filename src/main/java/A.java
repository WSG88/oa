import java.util.HashSet;

public class A {
    public static void main(String[] args) {
//        int angel = 30;
//        System.out.println(Math.PI * angel / 180);
//        System.out.println("\u2103");
//        System.out.println("\u2856");
        String s = "60°±30";
        String s1 = "≈150";
        String s2 = "153°±30′";
        String s3 = "Φ55±0.1";
        String s4 = "≯5";
        String s5 = "R5±1";
        HashSet hashSet = new HashSet();

        char[] ch = s2.toCharArray();
        for (char c : ch) {
            hashSet.add(c);
        }

        System.out.println(hashSet);

    }
}
