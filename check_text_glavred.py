from selenium import webdriver
from time import sleep
from openpyxl import load_workbook

class bot:
    def __init__(self):
        self.driver = webdriver.Firefox()
        self.get_texts()

    def get_texts(self):
        file = 'texts.xlsx'
        wb = load_workbook(file)
        sheet = wb.get_sheet_by_name('Лист2')
        row_count = sheet.max_row
        for x in range(2, row_count + 1):
            sleep(6)
            texts = sheet['A' + str(x)].value
            self.clean_text = texts.replace('_x000D_', '')
            self.navigate(self.clean_text)
            sheet['B' + str(x)].value = self.rating
            wb.save('texts.xlsx')
        self.driver.quit()

    def navigate(self, clean_text):
        url = 'https://glvrd.ru/'
        self.driver.get(url)
        form = self.driver.find_element_by_class_name('ql-editor')
        form.clear()
        form.send_keys(self.clean_text)
        sleep(5)
        self.rating = self.driver.find_element_by_xpath('//*[@id="app"]/div/div[3]/div[1]/div[2]/div[3]/div[1]/span[1]').text
        print(self.rating)
        return self.rating
        
def main():
    b = bot()


if __name__ == "__main__":
    main()