package com.tama.basicpoi;

import java.util.ArrayList;
import java.util.Iterator;

public class IteratorTest {

    // Membuat sebuah ArrayList
    private static final ArrayList<String> names = new ArrayList<>();

    public static void createList() {
        // Menambahkan elemen ke dalam ArrayList
        names.add("Alice");
        names.add("Bob");
        names.add("Charlie");
        names.add("Diana");

        showList();
    }

    public static void showList() {
        // Menggunakan Iterator untuk mengiterasi elemen-elemen dalam ArrayList
        Iterator<String> iterator = names.iterator();
        // Mengiterasi elemen-elemen menggunakan Iterator
        System.out.println("\nIterating over the list of names:");
        while (iterator.hasNext()) {
            String name = iterator.next();
            System.out.println(name);
        }
    }

    public static void removeItem() {
        // Menggunakan Iterator untuk mengiterasi elemen-elemen dalam ArrayList
        Iterator<String> iterator = names.iterator();
        // Mengiterasi elemen-elemen menggunakan Iterator dan menghapus elemen "Charlie"
        System.out.println("\nIterating and removing 'Charlie' from the list of names:");
        while (iterator.hasNext()) {
            String name = iterator.next();
            if (name.equals("Charlie")) {
                iterator.remove();
            }
        }

        showList();
    }
}
