import org.testng.annotations.Test;

import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.util.ArrayList;

public class database_printing extends JDBCParent{

    @Test
    public void test1() {
        ArrayList<ArrayList<String>> list =
                ExcelUtility.getListData("src/test/java/JDBC.xlsx", "Sheet1", 1);
        for (ArrayList<String> e : list) {
            getTable(e.get(0));
        }
    }

    public void getTable(String a) {
        try {
            ResultSet rs = statement.executeQuery(a);
            ResultSetMetaData rsmd = rs.getMetaData();

            for (int i = 1; i <= rsmd.getColumnCount(); i++) {
                System.out.printf("%-15s", rsmd.getColumnName(i));
            }
            System.out.println();

            while (rs.next()) {
                for (int i = 1; i <= rsmd.getColumnCount(); i++) {
                    System.out.printf("%-15s", rs.getString(i));
                }
                System.out.println();
            }
        } catch (SQLException e) {
            throw new RuntimeException(e);
        }
    }

}
