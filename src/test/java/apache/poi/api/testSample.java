package apache.poi.api;

import java.io.IOException;
import java.util.ArrayList;

public class testSample {
    public static void main(String[] args) throws IOException {
        RetrieveExcelData retrieveExcelData = new RetrieveExcelData();
        ArrayList<String> data = retrieveExcelData.getData("Purchase");
        System.out.println(data);
    }
}
