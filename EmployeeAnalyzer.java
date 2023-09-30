import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class EmployeeAnalyzer {
    public static void main(String[] args) {
        try {
            // Open the Excel file for reading (replace "input.xlsx" with your file's path)
            FileInputStream fis = new FileInputStream("C:\\chakri\\EmployerAnalysis\\input");
            Workbook workbook = new XSSFWorkbook(fis);

            // Assuming the data is in the first sheet (You can specify the sheet name or index)
            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                // Skip the header row if needed
                if (row.getRowNum() == 0) {
                    continue;
                }

                Cell positionIdCell = row.getCell(0);
                Cell positionStatusCell = row.getCell(1);
                Cell timeInCell = row.getCell(2);
                Cell timeOutCell = row.getCell(3);
                Cell timecardHoursCell = row.getCell(4);
                Cell startDateCell = row.getCell(5);
                Cell endDateCell = row.getCell(6);
                Cell employeeNameCell = row.getCell(7);
                Cell fileNumberCell = row.getCell(8);

                String positionId = positionIdCell.getStringCellValue();
                String positionStatus = positionStatusCell.getStringCellValue();
                String timeIn = timeInCell.getStringCellValue();
                String timeOut = timeOutCell.getStringCellValue();
                double timecardHours = timecardHoursCell.getNumericCellValue();
                String startDate = startDateCell.getStringCellValue();
                String endDate = endDateCell.getStringCellValue();
                String employeeName = employeeNameCell.getStringCellValue();
                String fileNumber = fileNumberCell.getStringCellValue();

                // Task 1: Employees who have worked for 7 consecutive days
                Cell employeeNameCell = currentRow.getCell(7); 

                // Extract the employee name from the cell
                String currentEmployeeName = employeeNameCell.getStringCellValue();
                

                // Check if this is the first record or if the employee name has changed
                if (lastEmployeeName == null || !lastEmployeeName.equals(currentEmployeeName)) {
                    consecutiveDaysCount = 1; // Reset consecutive days count for a new employee
                    lastEmployeeName = currentEmployeeName; // Update the last employee name
                } else {
                    consecutiveDaysCount++; // Increment consecutive days count
                }

                // Check if the employee has worked for 7 consecutive days
                if (consecutiveDaysCount == 7) {
                    found = true;
                } else if (consecutiveDaysCount > 7) {
                    // Reset the count if it exceeds 7 (e.g., if the employee works for more than 7 consecutive days)
                    consecutiveDaysCount = 1;
                }

                // Print the employee name and position if they have worked for 7 consecutive days
                if (found) {
                    Cell positionIDCell = currentRow.getCell(0); // Assuming the position ID is in the 1st column (index 0)
                    System.out.println("Employee Name: " + currentEmployeeName);
                    System.out.println("Position ID: " + positionIDCell.getStringCellValue());
                }
            }

                // Task 2: Employees with less than 10 hours between shifts but greater than 1 hour
                 Iterator<Row> iterator = sheet.iterator();

            SimpleDateFormat dateFormat = new SimpleDateFormat("HH:mm"); // Format for time (24-hour)

            Date previousEndTime = null; // To track the end time of the previous shift
            String lastEmployeeName = null; // To keep track of the previous employee

            while (iterator.hasNext()) {
                Row currentRow = iterator.next();
                Cell employeeNameCell = currentRow.getCell(7); // Assuming the employee name is in the 8th column (index 7)
                Cell startTimeCell = currentRow.getCell(4); // Assuming the start time is in the 5th column (index 4)

                // Extract the employee name and start time from the cells
                String currentEmployeeName = employeeNameCell.getStringCellValue();
                String startTimeStr = startTimeCell.getStringCellValue();

                Date startTime = dateFormat.parse(startTimeStr);

                // Check if this is the first record or if the employee name has changed
                if (lastEmployeeName == null || !lastEmployeeName.equals(currentEmployeeName)) {
                    previousEndTime = null; // Reset previous end time for a new employee
                    lastEmployeeName = currentEmployeeName; // Update the last employee name
                } else {
                    // Check the time gap between shifts
                    if (previousEndTime != null) {
                        long timeGap = startTime.getTime() - previousEndTime.getTime();
                        
                        // Check if the time gap is less than 10 hours but greater than 1 hour
                        if (timeGap < 10 * 60 * 60 * 1000 && timeGap > 1 * 60 * 60 * 1000) {
                            Cell positionIDCell = currentRow.getCell(0); // Assuming the position ID is in the 1st column (index 0)
                            System.out.println("Employee Name: " + currentEmployeeName);
                            System.out.println("Position ID: " + positionIDCell.getStringCellValue());
                        }
                    }
                }

                // Update the previous end time
                previousEndTime = startTime;
            }


                // Task 3: Employees who have worked for more than 14 hours in a single shift
                if (timecardHours > 14) {
                    System.out.println("Employee Name: " + employeeName);
                    System.out.println("Position ID: " + positionId);
                    
                }
            }

            // Close the Excel file
            workbook.close();
            fis.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
