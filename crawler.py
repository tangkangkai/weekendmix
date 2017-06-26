# -*- coding: UTF-8 -*-

from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.alert import Alert
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import xlsxwriter
from datetime import datetime
import copy


class ExcelRow(object):
    def __init__(self, startOdds, columnM, columnN, endOdds):
        self.end_odds = endOdds
        self.start_odds = startOdds
        self.column_m = columnM
        self.column_n = columnN

class Odds(object):
    def __init__(self, oddsList):
        self.odds_list = oddsList

class Element(object):
    def __init__(self, content, color):
        self.content = content
        self.color = color

def switchTab(driver, tabNum):
    while tabNum >= len(driver.window_handles):
        driver.implicitly_wait(1) # seconds
    driver.switch_to_window(driver.window_handles[tabNum])

def getOddsElement(rowElem, colNum, game_result, redcard_result):
    ''' only used in OddsHistory page'''
    path = "td[" + str(colNum) + "]"
    content = rowElem.find_element_by_xpath(path).text
    path += "/b/font"
    color = ""
    home_score = int(game_result[0])
    away_score = int(game_result[-1])
    red_color = "red"
    if 1 <= colNum <= 3: #Copy color from browser
        color = rowElem.find_element_by_xpath(path).get_attribute('color')
        if color:
            color = color.encode('ascii', 'ignore')
    elif (colNum == 4 and redcard_result[0]) or (colNum == 6 and redcard_result[-1]):
        color = red_color #set color according to red card
    elif (colNum == 8 and home_score > away_score) or (colNum == 9 and home_score == away_score) or (colNum == 10 and home_score < away_score):
        color = red_color #set color according to game result
    return Element(content, color)

def markTimeRedIfUpset(odds_list, game_result):
    home_odds = float(odds_list[0].content)
    away_odds = float(odds_list[2].content)
    home_score = int(game_result[0])
    away_score = int(game_result[-1])
    if (home_odds > away_odds and home_score > away_score) or (home_odds < away_odds and home_score < away_score):
        odds_list[-1].color = "red"


driver = webdriver.Firefox()
driver.get("http://zq.win007.com/cn/SubLeague/2.html")

year = int(driver.title.encode('ascii','ignore')[2:4])
contents = driver.find_elements_by_css_selector('td.lsm2') #round

default_solution_set = False
my_solution_name = 'Auto Solution'
#company_list = ['BINGOAL']
#company_list =['bet 365']
company_list =['竞彩官方']

excel_rows = []


for i in range(0, 29):
    driver.implicitly_wait(8) #wait to bypass the website restriction
    contents[i].click()
    
    column_m = year * 100 + i + 1
    column_n = 0

    tableElem = driver.find_element_by_id("Table3")
    games = tableElem.find_elements_by_css_selector("tr")[2:]
    
    for game in games: #game
        driver.implicitly_wait(8) #wait to bypass the website restriction
        column_n += 1

        game_result = game.find_element_by_xpath("td[4]").text
        if not game_result or '-' not in game_result:
            continue
        home_redcard = game.find_elements_by_xpath("td[3]/span")
        away_redcard = game.find_elements_by_xpath("td[5]/span")
        redcard_result = [home_redcard, away_redcard]

        link = game.find_element_by_link_text('[欧]')
        link.click()
        switchTab(driver, 1)
        if not default_solution_set: #custom solution
            #1 Hover
            menu_to_hover_over = WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, "a_solutions")))
            #driver.implicitly_wait(5)
            #menu_to_hover_over = driver.find_element_by_id("a_solutions")
            hover_action = ActionChains(driver).move_to_element(menu_to_hover_over)
            hover_action.perform() 
            
            #2 Click
            set_solution_button = driver.find_element_by_class_name("zdbtn")
            set_solution_button.click()

            #3 Select Company
            for company in company_list:
                company_to_select = driver.find_element_by_link_text(company)
                company_to_select.click()

            #4 Enter solution name
            solution_name_input = driver.find_element_by_id("solution_name")
            solution_name_input.clear()
            solution_name_input.send_keys(my_solution_name)

            #5. Save solution
            solution_save_button = driver.find_element_by_id("solution_saveBtn")
            solution_save_button.click()
            Alert(driver).accept()

            #6. Close pop up
            close_button = driver.find_element_by_css_selector("span[onclick*=company]");
            close_button.click()

            #7. Select custom solution
            hover_action.perform()
            driver.find_element_by_link_text(my_solution_name).click()

            default_solution_set = True
        
        #Open detail odds page
        try:
            row_button = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "td[onclick*=OddsHistory")))
        except:
            driver.close()
            switchTab(driver, 0)
            continue
        #row_button = driver.find_element_by_css_selector("td[onclick*=OddsHistory")
        row_button.click()
        switchTab(driver, 2)

        WebDriverWait(driver, 8).until(EC.presence_of_element_located((By.XPATH, "//tr")))
        changedTime = driver.find_elements_by_css_selector("font[color=blue]")
        count = 0
        while not changedTime:
            driver.implicitly_wait(10)
            changedTime = driver.find_elements_by_css_selector("font[color=blue]")
            count += 1
            if count > 5:
                raise Exception("fuck")

        odds_elements = driver.find_elements_by_xpath("//tr")
        end_odds_tr = odds_elements[1]
        start_odds_tr = odds_elements[-1]
        
        #Generate odds
        start_odds_list = []
        end_odds_list = []
        for column in range(1, 12): #11 columns
            end_odds_list += getOddsElement(end_odds_tr, column, game_result, redcard_result),
            start_odds_list += getOddsElement(start_odds_tr, column, game_result, redcard_result),
        #markTimeRedIfUpset(end_odds_list, game_result)
        #markTimeRedIfUpset(start_odds_list, game_result)
        
        start_odds = Odds(start_odds_list)
        end_odds = Odds(end_odds_list)
        excel_row = ExcelRow(start_odds, column_m, column_n, end_odds)
        excel_rows += excel_row,

        #Go back to home page
        driver.close()
        switchTab(driver, 1)
        driver.close()
        switchTab(driver, 0)

#-- Export to Excel--
def getDefaultFormat(workbook):
    format = workbook.add_format()
    format.set_font_name('Tahoma')
    format.set_font_size(9)
    format.set_border(2)
    format.set_border_color("gray")
    return format

file_name = str(datetime.now())[0:19] + ".xlsx"
workbook = xlsxwriter.Workbook(file_name)
worksheet = workbook.add_worksheet()

for row in range(len(excel_rows)):
    start_odds_list = excel_rows[row].start_odds.odds_list
    end_odds_list = excel_rows[row].end_odds.odds_list
    column = 0
    for j in range(len(end_odds_list)):
        odds = end_odds_list[j]
        format = getDefaultFormat(workbook)
        if odds.color == 'red':
            format.set_font_color('red')
        elif odds.color == 'green':
            format.set_font_color('green')
        if 0 <= j <= 2:
            format.set_bold(True)
        if j == 6:
            format.set_bg_color("silver")
        worksheet.write(row, column, odds.content, format)
        column += 1
    worksheet.write(row, column, excel_rows[row].column_m, getDefaultFormat(workbook))
    column += 1
    worksheet.write(row, column, excel_rows[row].column_n, getDefaultFormat(workbook))
    column += 1
    #write start odds
    for j in range(len(start_odds_list)):
        odds = start_odds_list[j]
        format = getDefaultFormat(workbook)
        if odds.color == 'green':
            format.set_font_color('green')
        if 0 <= j <= 2:
            format.set_bold(True)
        if j == 6:
            format.set_bg_color("silver")
        worksheet.write(row, column, odds.content, format)
        column += 1
workbook.close()

''' Documentation
onhover:
move_to_element
http://stackoverflow.com/questions/8252558/is-there-a-way-to-perform-a-mouseover-hover-over-an-element-using-selenium-and

input text:
send_keys
http://stackoverflow.com/questions/18557275/locating-entering-a-value-in-a-text-box-using-selenium-and-python

ignore unicode:
[code].encode('ascii','ignore')
'''


