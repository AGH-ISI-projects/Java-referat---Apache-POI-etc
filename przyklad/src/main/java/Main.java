import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

public class Main {

    private static final String[] KOLUMNY = {"Imie", "Stanowisko", "Pensja"};

    public static void main(String[] args) throws IOException {
        Path plikCsv = Path.of("dane.csv");
        Path plikXlsx = Path.of("wynik.xlsx");

        CSVFormat format = CSVFormat.DEFAULT.builder()
                .setDelimiter(';')
                .setHeader()
                .setIgnoreEmptyLines(true)
                .setTrim(true)
                .build();

        try (BufferedReader reader = Files.newBufferedReader(plikCsv, StandardCharsets.UTF_8)) {
            pominBom(reader);

            try (CSVParser parser = format.parse(reader);
                 Workbook workbook = new XSSFWorkbook()) {

                for (String kolumna : KOLUMNY) {
                    if (!parser.getHeaderMap().containsKey(kolumna)) {
                        throw new IllegalStateException("Brak kolumny w pliku CSV: " + kolumna);
                    }
                }

                Sheet sheet = workbook.createSheet("Pracownicy");

                Font pogrubiona = workbook.createFont();
                pogrubiona.setBold(true);
                CellStyle stylNaglowka = workbook.createCellStyle();
                stylNaglowka.setFont(pogrubiona);

                Row naglowek = sheet.createRow(0);
                for (int i = 0; i < KOLUMNY.length; i++) {
                    Cell komorka = naglowek.createCell(i);
                    komorka.setCellValue(KOLUMNY[i]);
                    komorka.setCellStyle(stylNaglowka);
                }

                int numerWiersza = 1;
                for (CSVRecord rekord : parser) {
                    if (!rekord.isConsistent()) {
                        throw new IllegalStateException("Zła liczba pól w rekordzie nr " + rekord.getRecordNumber());
                    }
                    BigDecimal pensja = new BigDecimal(rekord.get("Pensja").replace(',', '.'));

                    Row wiersz = sheet.createRow(numerWiersza++);
                    wiersz.createCell(0).setCellValue(rekord.get("Imie"));
                    wiersz.createCell(1).setCellValue(rekord.get("Stanowisko"));
                    wiersz.createCell(2).setCellValue(pensja.doubleValue()); // liczba, nie tekst
                }

                for (int i = 0; i < KOLUMNY.length; i++) {
                    sheet.setColumnWidth(i, 20 * 256);
                }
                try (OutputStream out = Files.newOutputStream(plikXlsx)) {
                    workbook.write(out);
                }

                System.out.println("Utworzono plik " + plikXlsx + ", wierszy danych: " + (numerWiersza - 1));
            }
        }
    }

    private static void pominBom(BufferedReader reader) throws IOException {
        reader.mark(1);
        if (reader.read() != '﻿') {
            reader.reset();
        }
    }
}
