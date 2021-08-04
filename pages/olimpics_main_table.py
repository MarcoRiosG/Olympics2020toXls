from selenium import webdriver
import xlwt

class MedalsTable():

    def __init__(self):

        self.driver = webdriver.Chrome()
        self.driver.maximize_window()
        self.driver.get("https://olympics.com/tokyo-2020/olympic-games/en/results/all-sports/medal-standings.htm")
        self.headers = []
        self.rows = []
        self.longest_name = ""

    def extract_table_headers(self):

        headers_row = self.driver.find_elements_by_xpath("//table[@id='medal-standing-table']/thead//th/div")
        medal_types = ["Gold", "Silver", "Bronze"]
        medal_index = 1

        for header in headers_row:
            if header.text:
                self.headers.append(header.text)
            elif medal_index == 1:
                self.headers.append(medal_types[medal_index - 1])
                medal_index += 1
            elif medal_index == 2:
                self.headers.append(medal_types[medal_index - 1])
                medal_index += 1
            elif medal_index == 3:
                self.headers.append(medal_types[medal_index - 1])
                medal_index += 1

        self.headers[-1] = self.headers[-1].replace("\n", " ")
        # print(self.headers)

    def extract_table_data(self):
        rows = self.driver.find_elements_by_xpath("//table[@id='medal-standing-table']/tbody/tr")

        for row in rows:
            self.rows.append(row.text.split())

        row_index = 0
        for country_row in self.rows:
            name_parts = []

            for country_data in country_row:
                if not country_data.isnumeric():
                    name_parts.append(country_data)

            name = " ".join(name_parts)
            if len(name) > len(self.longest_name):
                self.longest_name = name
            new_row = []
            cell_index = 0

            for country_data in country_row:
                if cell_index == 1:
                    new_row.append(name)
                elif not country_data.isnumeric():
                    None
                else:
                    new_row.append(country_data)
                cell_index += 1

            self.rows[row_index] = new_row
            row_index += 1

        # print(self.rows)
        # print(len(rows))

    def close_driver(self):
        self.driver.quit()

    def order_by_alphabet(self):
        self.rows = sorted(self.rows, key=lambda country_name: country_name[1])
        # print(self.rows)

    def add_to_xls(self):
        file_driver = xlwt.Workbook()
        top_countries = file_driver.add_sheet("Medallero (A-Z)")
        row_index = 0
        header_index = 0

        for header_name in self.headers:
            top_countries.write(row_index, header_index, header_name)
            top_countries.col(header_index).width = len(header_name) * 400
            header_index += 1

        row_index += 1

        for row in self.rows:
            column_index = 0

            for data in row:
                top_countries.write(row_index, column_index, data)
                column_index += 1

            row_index += 1

        top_countries.col(1).width = len(self.longest_name) * 256
        file_driver.save("countries_medals.xls")


medals = MedalsTable()
medals.extract_table_headers()
medals.extract_table_data()
medals.close_driver()
medals.order_by_alphabet()
medals.add_to_xls()
