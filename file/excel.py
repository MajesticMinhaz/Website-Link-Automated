import xlsxwriter


class ExcelCompanyNameToCompanyWebLink:
    def __init__(self, filename, company_name, company_web_url, date_time):
        self.filename = filename
        self.company_name = company_name
        self.company_web_url = company_web_url
        self.date_time = date_time

    def write_excel(self) -> str:
        workbook = xlsxwriter.Workbook(self.filename + ".xlsx")
        work_sheet = workbook.add_worksheet()
        work_sheet.write("A1", "Company Name")
        work_sheet.write("B1", "URL")
        work_sheet.write("C1", "Entry Date and Time")

        for value in range(len(self.company_name)):
            work_sheet.write(value + 1, 0, self.company_name[value])
            work_sheet.write(value + 1, 1, self.company_web_url[value])
            work_sheet.write(value + 1, 2, self.date_time[value])
        workbook.close()
        return f"You have successful to data Entry on \'{self.filename}.xlsx\' file.\n\nPlease Check It Now..."

