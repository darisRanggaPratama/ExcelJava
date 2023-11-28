package basic;

import java.util.Scanner;

public class Data {
    public static String fileXl, sheetXl;
    public static int beginRow, endRow, firstColumn, lastColumn;
    public static Scanner scan = new Scanner(System.in);

    public static void inputData() {
        print("\nInput\nFile name: ");
        fileXl = scan.nextLine();
        print("\nSheet name: ");
        sheetXl = scan.nextLine();
        print("\nFirst Row: ");
        beginRow = scan.nextInt();
        print("Last Row: ");
        endRow = scan.nextInt();
        print("First Column: ");
        firstColumn = scan.nextInt();
        print("Last Column: ");
        lastColumn = scan.nextInt();
    }

    public static void print(String str) {
        System.out.println(str);
    }
}
