package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashSet;
import java.util.Set;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;
import java.util.stream.StreamSupport;

public class ExcelToTimestampedFolderConverter {
    private static Set<String> filenameSet = new HashSet<>();

    public static void main(String[] args) {
        if (args.length < 1) {
            System.out.println("Usage: java -jar YourJarName.jar <ExcelFilePath>");
            System.exit(1);
        }
        String excelFilePath = args[0];
        String outputRootFolder = System.getProperty("user.dir");

        try (InputStream inputStream = new FileInputStream(new File(excelFilePath));
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            Sheet sheet = workbook.getSheetAt(0);
            // Extract the base name of the input file
            String inputFileName = new File(excelFilePath).getName();
            String baseInputFileName = inputFileName.substring(0, inputFileName.lastIndexOf('.'));

            // Create a timestamp for the folder name
            SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
            String timestamp = dateFormat.format(new Date());

            // Create the timestamped output folder
            String outputFolderPath = outputRootFolder + "\\" + baseInputFileName + "_" + timestamp;
            Files.createDirectories(Paths.get(outputFolderPath));

            AtomicInteger rowNumber = new AtomicInteger(1);

            StreamSupport.stream(sheet.spliterator(), false)
                    .skip(1) // Skip the header row
                    .forEach(row -> {
                        // Extract the first column as the file name
                        String fileName = cellToString(row.getCell(0));

                        // Validate filename to ensure no duplicates
                        if (!isValidFilename(fileName)) {
                            System.err.println("Error: Duplicate filename detected - " + fileName);
                            System.exit(1); // Exit program if duplicate found
                        }

                        // Add the filename to the set
                        filenameSet.add(fileName);

                        // Process the remaining columns as the content
                        String lineData = StreamSupport.stream(row.spliterator(), false)
                                .skip(1) // Skip the first column
                                .map(ExcelToTimestampedFolderConverter::cellToString)
                                .collect(Collectors.joining(","));

                        // Create the full path for the file
                        String filePath = outputFolderPath + "\\" + fileName;

                        try (BufferedWriter writer = new BufferedWriter(new FileWriter(new File(filePath)))) {
                            writer.write(lineData);
                            System.out.println("Conversion for row " + rowNumber.getAndIncrement() + " completed successfully.");
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    });

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String cellToString(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf((long) cell.getNumericCellValue());
            default:
                return "";
        }
    }

    private static boolean isValidFilename(String fileName) {
        return filenameSet.add(fileName);
    }
}
