import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileReader;
import java.io.FileOutputStream;
import java.io.Reader;

public class Main {
    public static void main(String[] args) throws Exception {

        CSVFormat format = CSVFormat.DEFAULT.builder()
                .setDelimiter(';')
                .setHeader()
                .build();

        try (Reader reader = new FileReader("dane.csv");
             CSVParser parser = format.parse(reader);
             XSSFWorkbook workbook = new XSSFWorkbook()) {

            XSSFSheet sheet = workbook.createSheet("Pracownicy");
            int numerWiersza = 0;

            for (CSVRecord rekordCsv : parser) {

                String imie = rekordCsv.get("Imie");
                String stanowisko = rekordCsv.get("Stanowisko");
                String pensja = rekordCsv.get("Pensja");

                XSSFRow wierszExcela = sheet.createRow(numerWiersza++);
                wierszExcela.createCell(0).setCellValue(imie);
                wierszExcela.createCell(1).setCellValue(stanowisko);
                wierszExcela.createCell(2).setCellValue(pensja);
            }

            try (FileOutputStream out = new FileOutputStream("wynik.xlsx")) {
                workbook.write(out);
            }

            System.out.println("Utworzono plik wynik.xlsx!");
        }
    }
}