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

    /**
     * Returns the HashMap of the Matrix in the Excel sheet.
     *  @param pExcelFileString the String path of the Excel sheet file
     * @param pSheetname the String of the Excel sheet workbook name
     * @return the HashMap of the Matrix.
     * @throws Exception
     */
    public static HashMap<String, HashMap<String, ValidationResultType>> getMatrix(final String pExcelFileString, final String pSheetname) throws Exception {

        HashMap<String, HashMap<String, ValidationResultType>> matrix =
                new HashMap<>();
        String path = null;

        URL resource = Thread.currentThread().getContextClassLoader().getResource(pExcelFileString);

        FileInputStream file = new FileInputStream(new File(resource.toURI()));

        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheet(pSheetname);

        // Iterate through each row one by one
        if (sheet == null) {
            workbook.close();
            file.close();
            throw new IOException("ERROR: Sheet for " + pSheetname + " + not available in the excel " + path);
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

    /**
     * Returns the {@link ValidationResultType} based on the contents of the cell String of the excel sheet
     * @param pCellValue the contents of the Excel cell
     * @return the {@link ValidationResultType}
     */
    private static ValidationResultType getValidationResult(final String pCellValue) {
        if (pCellValue.equalsIgnoreCase("*") || pCellValue.equalsIgnoreCase("ok"))
            return ValidationResultType.OK;
        else if (pCellValue.equalsIgnoreCase("b"))
            return ValidationResultType.ERROR;
        else if (pCellValue.equalsIgnoreCase("w"))
            return ValidationResultType.WARNING;
        else return ValidationResultType.ERROR;
        //If blank, then error, like in the standard transitions
    }

    /**
     * Parses the String of the Excel sheet based on its {@link org.apache.poi.ss.usermodel.CellType}
     * @param pCell the cell
     * @return the parsed String
     */
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
        HashMap<String, HashMap<String, ValidationResultType>> matrix = getMatrix("copy1.xlsx", "thirdparty");
    }
}
