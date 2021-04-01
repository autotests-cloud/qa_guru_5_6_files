package utils;

import com.codeborne.pdftest.PDF;
import com.codeborne.xlstest.XLS;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.Iterator;

public class Files {
    public static String readTextFromFile(File file) throws IOException {
        return FileUtils.readFileToString(file, StandardCharsets.UTF_8);
    }

    public static String readTextFromPath(String path) throws IOException {
        return readTextFromFile(getFile(path));
    }

    public static File getFile(String path) {
        return new File(path);
    }

    public static PDF getPdf(String path) throws IOException {
        return new PDF(getFile(path));
    }

    public static XLS getXls(String path) throws IOException {
        return new XLS(getFile(path));
    }

    public static String readXlsxFromPath(String path){
        StringBuilder sb = new StringBuilder();

        try (FileInputStream fis = new FileInputStream(path);
             Workbook myExcelBook = WorkbookFactory.create(fis)) {

            for (Sheet sheet : myExcelBook) {
                sb.append("Sheet ").append(sheet.getSheetName()).append(":\n");
                for (Row row : sheet) {
                    for (Cell cell : row) {
                        CellType cellType = cell.getCellType();
                        switch (cellType) {
                            case STRING:
                                sb.append(cell.getStringCellValue());
                                break;

                            case NUMERIC:
                                sb.append("[").append(cell.getNumericCellValue()).append("]");
                                break;

                            case FORMULA:
                                sb.append("{").append(cell.getCellFormula()).append(cell.getNumericCellValue()).append("}");

                                break;
                            default:
                                sb.append(cell.toString());
                                break;
                        }
                        sb.append(" ");
                    }
                    sb.append("\n");
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        return sb.toString();
    }
}
