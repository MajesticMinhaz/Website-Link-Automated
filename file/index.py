try:
    from googlesearch import search
    from excel import ExcelCompanyNameToCompanyWebLink
    import datetime
    import xlsxwriter
    import xlrd
    import os

    company_names = []
    company_web_urls = []
    currentDateTime = []
    user_input_file_name = input('Enter your Excel file name without extension (.xls) : ')
    user_input_sheet_index = int(input('Enter your Excel Sheet Number (1): '))
    full_file_name = f'{user_input_file_name}.xls'
    print(f'You Entered your file name as {full_file_name}')
    read_work_book = xlrd.open_workbook(full_file_name)
    read_work_sheet = read_work_book.sheet_by_index(user_input_sheet_index - 1)
    for number_of_rows in range(1, read_work_sheet.nrows):
        company_names.append(read_work_sheet.cell_value(number_of_rows, 0))
    for one_company_name in company_names:
        query = one_company_name
        for link in search(query, num_results=1, lang='en'):
            weblink = link
            dateTime = datetime.datetime.now().strftime("%I:%M%p on %B %d, %y")
            company_web_urls.append(weblink)
            currentDateTime.append(dateTime)

    writeExcel = ExcelCompanyNameToCompanyWebLink(user_input_file_name, company_names, company_web_urls, currentDateTime)
    result = writeExcel.write_excel()
    print(result)
except ImportError:
    print('One module is missing. please check and run again')
