import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Array;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

/**
 * Author: Harry Walker
 * <p>
 * Descrption...
 */
public class MatrixTest {

    enum ValidationResultType {
        OK,
        WARNING,
        ERROR,
        NA
    }

    public static HashMap<String, HashMap<String, ValidationResultType>> getMatrix() throws Exception {

        HashMap<String, HashMap<String, ValidationResultType>> matrix =
                new HashMap<>();
        String path = null;

        URL resource = Thread.currentThread().getContextClassLoader().getResource("copy1.xlsx");
        try {
            path = "C://Users/Harry/Documents/Wise/copy1.xlsx";
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

        FileInputStream file = new FileInputStream(new File(resource.toURI()));

        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheet("thirdparty");

        // Iterate through each row one by one
        if (sheet == null) {
            System.out.println("ERROR: Sheet for Demotion not available in the excel " + path);
            workbook.close();
            file.close();
            throw new IOException("ERROR: Sheet for Demotion not available in the excel" + path);
        }


        List<String> mlcFrom = new ArrayList<>();
        List<String> mlcTo = new ArrayList<>();

        Iterator<Row> rowIterator = sheet.rowIterator();

        while (rowIterator.hasNext()) {

            Row row = rowIterator.next();

            String mlcStatusFrom = "";
            if (row.cellIterator().hasNext()) {
                if (row.cellIterator().next().getRowIndex() != 0) {
                    mlcStatusFrom = parseCell(row.cellIterator().next());
                } else {
                    mlcStatusFrom = "10";
                }
            }
            Iterator<Cell> cellIterator = row.cellIterator();

            int i = 0;
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();

                String cellValue = parseCell(cell);

                System.out.println(cellValue);
            }


        }

        return matrix;
    }

    private static ValidationResultType getValidationResult(final String pCellValue) {
        if (pCellValue.equalsIgnoreCase("*") || pCellValue.equalsIgnoreCase("ok"))
            return ValidationResultType.OK;
        else if (pCellValue.equalsIgnoreCase("b"))
            return ValidationResultType.ERROR;
        else if (pCellValue.equalsIgnoreCase("w"))
            return ValidationResultType.WARNING;
        else return ValidationResultType.NA;
    }

    private static String parseCell(final Cell pCell) {
        switch (pCell.getCellType()) {
            case NUMERIC:
                return String.valueOf(pCell.getNumericCellValue());
            case STRING:
                return String.valueOf(pCell.getStringCellValue());
            default:
                return "";
        }
    }

    public static void main(String[] args) throws Exception {
        HashMap<String, HashMap<String, ValidationResultType>> matrix = getMatrix();
    }
}
