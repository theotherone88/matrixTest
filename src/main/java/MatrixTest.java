import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.net.URL;
import java.util.*;

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
    }

    /**
     * Returns the HashMap of the Matrix in the Excel sheet.
     *  @param pExcelFileString the String path of the Excel sheet file
     * @param pSheetname the String of the Excel sheet workbook name
     * @return the HashMap of the Matrix.
     * @throws Exception
     */
    public static HashMap<String, LinkedHashMap<String, ValidationResultType>> getMatrix(final String pExcelFileString,
                                                                                   final String pSheetname)
            throws Exception {

        HashMap<String, LinkedHashMap<String, ValidationResultType>> matrix =
                new HashMap<>();
        String path = null;

        URL resource = Thread.currentThread().getContextClassLoader().getResource(pExcelFileString);

        final FileInputStream file = new FileInputStream(new File(resource.toURI()));

        final XSSFWorkbook workbook = new XSSFWorkbook(file);
        final XSSFSheet sheet = workbook.getSheet(pSheetname);

        // Iterate through each row one by one
        if (sheet == null) {
            workbook.close();
            file.close();
            throw new IOException("ERROR: Sheet for " + pSheetname + " + not available in the excel " + path);
        }

        List<String> mlcTo = new ArrayList<>();

        LinkedHashMap<String, Integer> mlcFromMap = new LinkedHashMap<>();

        Iterator<Row> rowIterator = sheet.rowIterator();

        while (rowIterator.hasNext()) {
            final Row row = rowIterator.next();
            final Iterator<Cell> cellIterator = row.cellIterator();

            int i = 0;
            while (cellIterator.hasNext()) {
                final Cell cell = cellIterator.next();

                final String cellValue = parseCell(cell);

                if (cell.getRowIndex() > 0 && cell.getColumnIndex() == 0) {
                    mlcFromMap.put(cellValue, cell.getRowIndex());
                }
                if (cell.getRowIndex() == 0 && cell.getColumnIndex() > 0) {
                    mlcTo.add(cellValue);
                }
            }
        }


        Map<String, Integer> sortedMap = sortByValue(mlcFromMap);

        for (Map.Entry<String, Integer> entry : sortedMap.entrySet()) {
            final Row innerRow = sheet.getRow(entry.getValue());
            Iterator<Cell> cellIterator = innerRow.cellIterator();

            LinkedHashMap<String, ValidationResultType> innerMap = new LinkedHashMap<>();
            int mlcToIndex = 0;
            while (mlcToIndex < mlcTo.size()) {
                if (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    if (cell.getColumnIndex() > 0 && cell.getRowIndex() > 0) {
                        final ValidationResultType result = getValidationResult(parseCell(cell));
                        final String mlcToStr = mlcTo.get(mlcToIndex);
                        innerMap.put(mlcToStr, result);

                        matrix.put(entry.getKey(), innerMap);

                        mlcToIndex++;
                    }
                }

            }
        }
        return matrix;
    }

    public static <K, V extends Comparable<? super V>> Map<K, V> sortByValue(Map<K, V> map) {
        List<Map.Entry<K, V>> list = new ArrayList<>(map.entrySet());
        list.sort(Map.Entry.comparingByValue());

        Map<K, V> result = new LinkedHashMap<>();
        for (Map.Entry<K, V> entry : list) {
            result.put(entry.getKey(), entry.getValue());
        }

        return result;
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
        //If blank, then error, like in the standard transitions, or NA in the other Excel sheets
    }

    /**
     * Parses the String of the Excel sheet based on its {@link org.apache.poi.ss.usermodel.CellType}
     * @param pCell the cell
     * @return the parsed String
     */
    private static String parseCell(final Cell pCell) {
        switch (pCell.getCellType()) {
            case NUMERIC:
                return String.valueOf((int)pCell.getNumericCellValue());
            case STRING:
                return String.valueOf(pCell.getStringCellValue());
            default:
                return "";
        }
    }

    public static void main(String[] args) throws Exception {
        HashMap<String, LinkedHashMap<String, ValidationResultType>> matrix = getMatrix("copy1.xlsx", "buy");

        int i = 0;
    }
}
