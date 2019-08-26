from selenium import webdriver
import random
from time import sleep
from openpyxl import load_workbook

class bot:
    def __init__(self):
        self.driver = webdriver.Chrome()
        self.sign_in()

    def sign_in(self):
        url = 'https://turgenev.ashmanov.com/'
        self.driver.maximize_window()
        self.driver.get(url)
        sign_in_form = self.driver.find_element_by_css_selector('body > div.topblock > div.mb.large > div.last > span')
        sign_in_form.click()
        login = self.driver.find_element_by_xpath('//*[@id="dlg_main_block"]/form/div[1]/div[1]/input')
        password = self.driver.find_element_by_xpath('//*[@id="dlg_main_block"]/form/div[1]/div[2]/input')
        login.send_keys('') # your login
        password.send_keys('') # your password
        button_sign_in = self.driver.find_element_by_xpath('//*[@id="dlg_main_block"]/form/div[2]/input')
        sleep(random.uniform(3, 4))
        button_sign_in.click()
        self.editor_activation()

    def editor_activation(self):
        sleep(random.uniform(3, 4))
        button_editor = self.driver.find_element_by_id('show_tb_btn')
        button_editor.click()
        self.get_texts()

    def get_texts(self):
        file = 'texts.xlsx'
        wb = load_workbook(file)
        sheet = wb.get_sheet_by_name('Лист2')
        row_count = sheet.max_row
        for x in range(2, row_count + 1):
            sleep(random.uniform(5, 7))
            texts = sheet['A' + str(x)].value
            self.clean_text = texts.replace('_x000D_', '')
            self.navigate(self.clean_text)
            sheet['C' + str(x)].value = self.rating
            sheet['D' + str(x)].value = self.words_count + ' слов'
            sheet['E' + str(x)].value = self.text_size + ' символов'
            wb.save('texts.xlsx')
        self.driver.quit()
    
    def navigate(self, clean_text):
        self.driver.switch_to.frame('textfield_ifr')
        form = self.driver.find_element_by_xpath('//*[@id="tinymce"]')
        form.clear()
        form.send_keys(self.clean_text)
        self.driver.switch_to.default_content()
        button_check_text = self.driver.find_element_by_xpath('//*[@id="recheck_btn"]')
        button_check_text.click()
        sleep(random.uniform(3, 4))
        self.rating = self.driver.find_element_by_xpath('//*[@id="infoblock"]/h2/span/span[1]/span').text
        self.words_count = self.driver.find_element_by_id('words_count').text
        self.text_size = self.driver.find_element_by_id('text_size').text
        print(self.rating, self.words_count, self.text_size)
        return self.rating, self.words_count, self.text_size

def main():
    b = bot()

if __name__ == "__main__":
    main()
