import java.util.ArrayList;
import java.util.List;

public class Cal {
    static void findNum(List<Integer> Arr) {
        int arr = Arr.get(0);
        for (int i = 0; i < Arr.size(); i++) {
            int val = Arr.get(i);
            if (arr < val) {
                arr = val;
            }
        }
        System.out.println("max = " + arr);
    }
    public static void main(String[] args) {
        List<Integer> Arr = new ArrayList<>(List.of(new Integer[]{1, 2, 1, 0, -1, 4}));
        findNum(Arr);
    }
}
