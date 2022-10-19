import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.util.ArrayList;

public class print_to_new_excel extends JDBCParent{
    //print all data in database actor table to new excel
    @Test
    public void test1() {
        ArrayList<String> rsList = new ArrayList<>();
        ArrayList<String> rsmdList = new ArrayList<>();
        try {
            ResultSet rs = statement.executeQuery("select * from actor");
            ResultSetMetaData rsmd = rs.getMetaData();
            for (int i = 1; i <= rsmd.getColumnCount(); i++) {
                rsmdList.add(rsmd.getColumnName(i));
            }
            while (rs.next()) {
                for (int i = 1; i <= rsmd.getColumnCount(); i++) {
                    rsList.add(rs.getString(i));
                }
            }
            rs.last();
            writeExcel(rs.getRow(), rsmd.getColumnCount(), rsList, rsmdList);
        } catch (SQLException e) {
            throw new RuntimeException(e);
        }

    }

    public void writeExcel(int b, int a, ArrayList<String> rsList, ArrayList<String> rsmdList) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sheet1");

        Row row = sheet.createRow(0);
        int j = 0;
        for (String e : rsmdList) {
            row.createCell(j).setCellValue(e);
            j++;
        }
        j = 1;
        int count = 0;

        for (int i = 0; i < b; i++) {
            row = sheet.createRow(j);
            for (int k = 0; k < a; k++) {
                row.createCell(k).setCellValue(rsList.get(count));
                count++;
            }
            j++;
        }

        try {
            FileOutputStream fileOutputStream = new FileOutputStream("src/test/java/jdbc2.xlsx");
            workbook.write(fileOutputStream);
            fileOutputStream.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
