#include <iostream>
#include <vector>
#include <xlsxio_read.h>

struct EmployeeShift {
    std::string name;
    std::string position;
    int consecutiveDays;
    int lessThan10Hours;
    int moreThan14Hours;
};

void analyzeExcel(const char* filePath, std::vector<EmployeeShift>& results) {
    xlsxioreader xlsioReader;
    xlsxioreadersheet xlsioSheet;

    xlsioReader = xlsxioread_open_file(filePath);
    if (!xlsioReader) {
        std::cerr << "Error opening Excel file" << std::endl;
        return;
    }

    xlsioSheet = xlsxioread_sheet_open(xlsioReader, NULL, XLSXIOREAD_SKIP_NONE);
    if (!xlsioSheet) {
        std::cerr << "Error opening Excel sheet" << std::endl;
        xlsxioread_close(xlsioReader);
        return;
    }

    int row = 0;
    std::string prevName;
    std::string prevDate;
    int consecutiveDays = 0;
    int lessThan10Hours = 0;
    int moreThan14Hours = 0;

    while (xlsxioread_sheet_next_row(xlsioSheet)) {
        const char* name = xlsxioread_sheet_next_cell(xlsioSheet);
        const char* position = xlsxioread_sheet_next_cell(xlsioSheet);
        const char* date = xlsxioread_sheet_next_cell(xlsioSheet);
        const char* startTime = xlsxioread_sheet_next_cell(xlsioSheet);
        const char* endTime = xlsxioread_sheet_next_cell(xlsioSheet);

        if (name && position && date && startTime && endTime) {
            if (prevName == name && prevDate != date) {
                consecutiveDays++;
                int hoursDiff = std::stoi(startTime) - std::stoi(prevEndTime);
                if (hoursDiff > 1 && hoursDiff < 10) {
                    lessThan10Hours++;
                }
                int shiftHours = std::stoi(endTime) - std::stoi(startTime);
                if (shiftHours > 14) {
                    moreThan14Hours++;
                }
            } else {
                consecutiveDays = 1;
                lessThan10Hours = 0;
                moreThan14Hours = 0;
            }

            if (consecutiveDays >= 7) {
                EmployeeShift result;
                result.name = name;
                result.position = position;
                result.consecutiveDays = consecutiveDays;
                result.lessThan10Hours = lessThan10Hours;
                result.moreThan14Hours = moreThan14Hours;
                results.push_back(result);
            }

            prevName = name;
            prevDate = date;
            prevEndTime = endTime;
        }
    }

    xlsxioread_sheet_close(xlsioSheet);
    xlsxioread_close(xlsioReader);
}

int main() {
    std::string excelFilePath = "E:\\new assgin\\Assignment_Timecard.xlsx";


    std::vector<EmployeeShift> results;

    analyzeExcel(excelFilePath, results);

    // Print the results to the console
    for (const auto& result : results) {
        std::cout << "Name: " << result.name << ", Position: " << result.position << std::endl;
        std::cout << "Consecutive Days: " << result.consecutiveDays << std::endl;
        std::cout << "Less Than 10 Hours Between Shifts: " << result.lessThan10Hours << std::endl;
        std::cout << "More Than 14 Hours in a Single Shift: " << result.moreThan14Hours << std::endl;
        std::cout << "------------------------" << std::endl;
    }

    return 0;
}
