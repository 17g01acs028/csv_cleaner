package org.example;

import org.apache.log4j.BasicConfigurator;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.*;
import java.util.concurrent.TimeUnit;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
public class Main {
    public static void main(String[] args) {
        BasicConfigurator.configure();
        long startTime = System.currentTimeMillis();
            String csvFilePath = "data.csv";
            String excelFilePath = "src/output.xlsx";

            try {
                BufferedReader bufferedReader = new BufferedReader(new FileReader(csvFilePath));

                // Create Excel workbook and sheets
                Workbook workbook = new SXSSFWorkbook();
                Sheet male = workbook.createSheet("male");
                Sheet female = workbook.createSheet("female");
                Sheet dataWithErrorSheet = workbook.createSheet("data_with_error");

                // Read the file line by line
                int count =1;
                String line;
                System.out.println("\uD83E\uDDF9 Cleaning up, filtering out the mess, and making your data sparkle! ✨");
                while ((line = bufferedReader.readLine()) != null) {

                    if(count > 1){
                        // Split the line into columns using a comma as the delimiter
                        String[] columns = line.split(",");
                        // Validate and clean the data
                        columns[2] = columns[2].replaceAll("\\s","");
                        columns[3] = columns[3].trim();
                        String errors = validateAndCleanData(columns);

                        // Determine the sheet based on gender
                        Sheet currentSheet = (columns[4].equalsIgnoreCase("male")) ? male :
                                (columns[4].equalsIgnoreCase("female")) ? female :
                                        dataWithErrorSheet;

                        if (currentSheet.getPhysicalNumberOfRows() == 0) {
                            boolean error;
                            error = !columns[4].equalsIgnoreCase("male") && !columns[4].equalsIgnoreCase("female");
                            addHeaderRow(currentSheet,error);
                        }

                        // Create a row and add data to the appropriate sheet
                        Row row = (errors.isEmpty()) ? currentSheet.createRow(currentSheet.getPhysicalNumberOfRows()) :
                                dataWithErrorSheet.createRow(dataWithErrorSheet.getPhysicalNumberOfRows());

                        for (int i = 0; i < columns.length; i++) {
                            Cell cell = row.createCell(i);
                            cell.setCellValue(columns[i]);
                        }


                        // Add an extra column for errors
                        Cell errorCell = row.createCell(columns.length);
                        errorCell.setCellValue(errors);
                    }

                    count++;
                }


                System.out.println("\uD83D\uDE80 Buckle up! We're launching your squeaky-clean data to the drive. Hold tight! \uD83C\uDF0C");

                // Save the workbook to a file
                try (FileOutputStream outputStream = new FileOutputStream(excelFilePath)) {
                    workbook.write(outputStream);
                }
                long endTime = System.currentTimeMillis(); // Capture the end time
                long executionTime = endTime - startTime;

                System.out.println("\uD83C\uDF89 Your Excel file has been saved successfully ✨");
                System.out.println("\uD83C\uDF89 Ta-da! We did it! Your data is now shining bright! \uD83C\uDF1F We've officially saved the day! \uD83D\uDE80 Time taken: "+formatDuration(executionTime)+". Thanks for hanging tight! ⏰ " );
                // Close the BufferedReader
                bufferedReader.close();

            } catch (Exception e) {
                 e.printStackTrace();
            }
    }

    private static void addHeaderRow(Sheet sheet, boolean error) {
        Row headerRow = sheet.createRow(0);

        // Add headers: ID, NAME, PHONE, EMAIL, GENDER
        String[] headers = (error) ? new String[]{"ID", "NAME", "PHONE", "EMAIL", "GENDER", "ERRORS"} :
                new String[]{"ID", "NAME", "PHONE", "EMAIL", "GENDER"};
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
        }
    }

    private static String formatDuration(long duration) {
        long hours = TimeUnit.MILLISECONDS.toHours(duration);
        long minutes = TimeUnit.MILLISECONDS.toMinutes(duration) % 60;
        long seconds = TimeUnit.MILLISECONDS.toSeconds(duration) % 60;
        long millis = duration % 1000;

        return String.format("%02d:%02d:%02d.%03d", hours, minutes, seconds, millis);
    }
    private static String validateAndCleanData(String[] data) {
        // Field validation
        String id = data[0];
        String phone = data[2];
        String email = data[3];
        String gender = data[4];

        StringBuilder errors = new StringBuilder();

        // Check ID
        if (!(id.length() == 8 && id.matches("\\d+"))) {
            errors.append("Invalid ID, ");
        }

        // Check Phone
        if (!(isValidMobileNo(phone))) {
            errors.append("Invalid Phone, ");
        }

        // Check Email
        if (!isValidEmail(email)) {
            errors.append("Invalid Email, ");
        }

        // Check Gender
        if (!(gender.equalsIgnoreCase("male") || gender.equalsIgnoreCase("female"))) {
            errors.append("Invalid Gender, ");
        }

        return errors.toString().replaceAll(", $", ""); // Remove trailing comma and space
    }
//Check if email is valid    
    private static boolean isValidEmail(String email) {
        // Use a simple regex for email validation
        String regex = "^[\\w.-]+@[\\w.-]+\\.[a-z]{2,}$";
        Pattern pattern = Pattern.compile(regex);
        Matcher matcher = pattern.matcher(email);
        return matcher.matches();
    }

//check if mobile Number is OK.
    private static boolean isValidMobileNo(String phoneNumber) {
        String regex = "^(254|0)([17][0-9]|[1][0-1]){1}[0-9]{1}[0-9]{6}$";
        return phoneNumber.matches(regex);
    }
}
