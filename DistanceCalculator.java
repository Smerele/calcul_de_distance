package com.example;

import com.google.maps.GeoApiContext;
import com.google.maps.DistanceMatrixApi;
import com.google.maps.errors.ApiException;
import com.google.maps.model.DistanceMatrix;
import com.google.maps.model.DistanceMatrixElement;
import com.google.maps.model.DistanceMatrixRow;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;

public class DistanceCalculator {

    public static void main(String[] args) {
        if (args.length < 1) {
            System.out.println("Utilisation: java -jar DistanceCalculator.jar <input_file>");
            return;
        }
        
        String inputFilePath = args[0];
        boolean isCsv = inputFilePath.endsWith(".csv");
        boolean isXlsx = inputFilePath.endsWith(".xlsx");
        if (isCsv) {
            inputFilePath = csvtoxlsx(inputFilePath);
        }
        
        String outputFilePath = "sortie_" + inputFilePath;

        GeoApiContext context = new GeoApiContext.Builder()
            .apiKey("YOUR GOOGLE API KEY")
            .build();

        try (FileInputStream fis = new FileInputStream(new File(inputFilePath));
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue;
                }

                Cell refCell = row.getCell(0);
                Cell addr1DepartCell = row.getCell(1);
                Cell addr2DepartCell = row.getCell(2);
                Cell codePostalDepartCell = row.getCell(3);
                Cell villeDepartCell = row.getCell(4);
                Cell inseeDepartCell = row.getCell(5);
                Cell inseeArriveeCell = row.getCell(6);
                Cell addr1ArriveeCell = row.getCell(7);
                Cell addr2ArriveeCell = row.getCell(8);
                Cell codePostalArriveeCell = row.getCell(9);
                Cell villeArriveeCell = row.getCell(10);
                Cell distanceCell = row.createCell(11);

                String origin = buildAddress(addr1DepartCell, addr2DepartCell, codePostalDepartCell, villeDepartCell);
                String destination = buildAddress(addr1ArriveeCell, addr2ArriveeCell, codePostalArriveeCell, villeArriveeCell);

                if (origin.isEmpty() || destination.isEmpty()) {
                    distanceCell.setCellValue("Adresse incomplete");
                    continue;
                }

                try {
                    DistanceMatrix result = DistanceMatrixApi.getDistanceMatrix(context, new String[]{origin}, new String[]{destination}).await();
                    for (DistanceMatrixRow matrixRow : result.rows) {
                        for (DistanceMatrixElement element : matrixRow.elements) {
                            if (element.distance != null) {
                                double distanceInKilometers = element.distance.inMeters / 1000.0;
                                distanceCell.setCellValue(distanceInKilometers + " km");
                                System.out.println("Distance de " + origin + " a " + destination + " est de " + distanceInKilometers + " km");
                            } else {
                                distanceCell.setCellValue("non trouvee");
                                System.out.println("Aucune distance trouvee de " + origin + " a " + destination);
                            }
                        }
                    }
                } catch (ApiException | InterruptedException | IOException e) {
                    e.printStackTrace();
                    distanceCell.setCellValue("Error");
                    System.out.println("Erreur de calcul de distance de " + origin + " a " + destination);
                }
            }

            try (FileOutputStream fos = new FileOutputStream(new File(outputFilePath))) {
                workbook.write(fos);
            }

            System.out.println("fichier cree a " + outputFilePath);

            if (isCsv || isXlsx) {
                String csvOutputFilePath = outputFilePath.replace(".xlsx", ".csv");
                xlsxtocsv(outputFilePath, csvOutputFilePath);
                System.out.println("Le fichier " + outputFilePath + " a été reconverti en CSV.");
            }

        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (isCsv || isXlsx) {
            	String csvOutputFilePath = outputFilePath.replace(".xlsx", ".csv");
            	xlsxtocsv(outputFilePath, csvOutputFilePath);
            	System.out.println("Le fichier " + outputFilePath + " a été reconverti en CSV.");
                File xlsxFile = new File(outputFilePath);
                if (xlsxFile.delete()) {
                    System.out.println("Le fichier XLSX temporaire " + outputFilePath + " a été supprimé.");
                } else {
                    System.out.println("Erreur lors de la suppression du fichier XLSX temporaire " + outputFilePath);
                }
            }
        }
    }

    private static String buildAddress(Cell addr1, Cell addr2, Cell postalCode, Cell city) {
        StringBuilder address = new StringBuilder();
        if (addr1 != null && addr1.getCellType() == CellType.STRING) {
            address.append(addr1.getStringCellValue()).append(", ");
        }
        if (addr2 != null && addr2.getCellType() == CellType.STRING) {
            address.append(addr2.getStringCellValue()).append(", ");
        }
        if (postalCode != null && postalCode.getCellType() == CellType.STRING) {
            address.append(postalCode.getStringCellValue()).append(" ");
        }
        if (city != null && city.getCellType() == CellType.STRING) {
            address.append(city.getStringCellValue());
        }
        return address.toString().trim();
    }

    private static String csvtoxlsx(String csvfile) {
        String xlsxfile = csvfile.replace(".csv", ".xlsx");
        try {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("sheet1");
            XSSFFont xssfFont = workbook.createFont();
            xssfFont.setCharSet(XSSFFont.ANSI_CHARSET);
            XSSFCellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setFont(xssfFont);
            String currentLine;
            int RowNum = 0;
            BufferedReader br = new BufferedReader(new FileReader(csvfile));
            while ((currentLine = br.readLine()) != null) {
                String[] str = currentLine.split(";");
                XSSFRow currentRow = sheet.createRow(RowNum++);
                for (int i = 0; i < str.length; i++) {
                    str[i] = str[i].replaceAll("\"", "");
                    str[i] = str[i].replaceAll("=", "");
                    XSSFCell cell = currentRow.createCell(i);
                    cell.setCellStyle(cellStyle);
                    cell.setCellValue(str[i].trim());
                }
            }
            FileOutputStream fileOutputStream = new FileOutputStream(xlsxfile);
            workbook.write(fileOutputStream);
            fileOutputStream.close();
            System.out.println("Conversion reussie : " + csvfile + " en " + xlsxfile);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return xlsxfile;
    }

    private static void xlsxtocsv(String xlsxfile, String csvfile) {
        try (FileInputStream fis = new FileInputStream(new File(xlsxfile));
             Workbook workbook = new XSSFWorkbook(fis);
             FileWriter csvWriter = new FileWriter(new File(csvfile))) {

            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                for (Cell cell : row) {
                    switch (cell.getCellType()) {
                        case STRING:
                            csvWriter.append(cell.getStringCellValue());
                            break;
                        case NUMERIC:
                            csvWriter.append(String.valueOf(cell.getNumericCellValue()));
                            break;
                        default:
                            csvWriter.append("");
                    }
                    csvWriter.append(";");
                }
                csvWriter.append("\n");
            }
            System.out.println("Conversion reussie : " + xlsxfile + " en " + csvfile);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
