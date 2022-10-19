import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;

public class insert_into_database_from_excel {
    @Test
    public void test1() {
        String path = "src/test/java/JDBC3.xlsx";
        Workbook workbook = null;
        try {
            FileInputStream inputStream = new FileInputStream(path);
            workbook = WorkbookFactory.create(inputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        Sheet sheet = workbook.getSheet("Sheet1");
        for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
            Row row = sheet.getRow(i);
            String a = "INSERT INTO jdbc (id,ad,soyad,telefon) " + "VALUES ('" + row.getCell(0) + "', '" + row.getCell(1).toString() + "','" +
                    row.getCell(2).toString() + "', '" + row.getCell(3).toString() + "')";
            writeDB("z_burak1", a);
        }
    }


    public void writeDB(String dbName, String a) {
        try {
            String url = "jdbc:mysql://db-........us-east-1.rds.amazonaws.com:.../" + dbName + "";
            String username = "...";
            String password = "...";
            Connection connection = DriverManager.getConnection(url, username, password);
            Statement statement = connection.createStatement();
            statement.executeUpdate(a);
            connection.close();
        } catch (SQLException e) {
            throw new RuntimeException(e);
        }

    }
}
