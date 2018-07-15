import classes.Reader;

import java.io.File;
import java.io.IOException;

public class main {

    public static void main(String[] args) throws IOException {
        File file = new File("List.xls");
        if (file.exists() && file.delete()) {
            System.out.println("Удален");
        } else {
            System.out.println("Не удален");
        }

        String path = "C:/Users/Student/Desktop/11.txt";
        Reader reader = new Reader(path);
        reader.ReadFile();
    }
}
