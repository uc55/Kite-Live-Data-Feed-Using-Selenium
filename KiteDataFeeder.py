import time
import datetime

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
from pandas import ExcelWriter
# from bs4 import BeautifulSoup
# import pdb


class AutoTrade(object):

    # Initialize the chrome driver
    def __init__(self):
        self.strategyExecuted = 0
        self.nifty50Quotes = pd.DataFrame()
        self.timeout = 5
        self.driver = None
        self.opt = Options()
        self.opt.add_experimental_option("debuggerAddress", "localhost:9999")

        self.service = Service('./lib/chromedriver.exe')
        self.service.start()
        self.driver = webdriver.Remote(self.service.service_url, options=self.opt)

    # To make sure we wait till the element appears
    def getCssElement(self, cssSelector):
        return WebDriverWait(self.driver, self.timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, cssSelector)))

    def getCssElementList(self, cssSelector):
        return WebDriverWait(self.driver, self.timeout).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, cssSelector)))

    def getChildElement(self, parentElement, by, cssAddress):
        return WebDriverWait(parentElement, self.timeout).until(
            EC.presence_of_element_located((by, cssAddress)))

    def getChildElementList(self, parentElement, by, cssAddress):
        return WebDriverWait(parentElement, self.timeout).until(
            EC.presence_of_all_elements_located((by, cssAddress)))


    def automation(self):
        # self.driver.get('https://kite.zerodha.com/')
        print('Started Data Feeder...')

        while True:
            start = time.time()
            self.watchlist()
            finish = time.time()
            print('Data Refresh Time -> ' + str(round(finish-start, 2)) + ' Sec.')
            # Bed Rest For 1 Second.
            time.sleep(1)

        # pdb.set_trace()
        # close chrome
        self.driver.quit()
        # self.service.stop()


    # Watchlist ====================================================================
    def watchlist(self):
        symbolL = []
        openL = []
        prevCloseL = []
        highL = []
        lowL = []
        closeL = []
        bidPriceL = []
        bidQtyL = []
        askPriceL = []
        askQtyL = []

        try:
            # Open 'Watchlist 5' Tab.
            link = self.getCssElement("li[data-balloon='Watchlist 5']")
            isSelected = link.get_attribute('class')
            if isSelected != 'item selected':
                link.click()

            # Now, load all parent elements from 'Watchlist 5' tab.
            for symbolInfoContainer in self.getCssElementList("div[draggable='true']"):
                css_class = symbolInfoContainer.get_attribute('class')

                # Open all the closed 'Market Depth window' to fill OHLC & Market Depth data
                if len(css_class) < 54:
                    self.driver.execute_script("arguments[0].scrollIntoView();", symbolInfoContainer)
                    hover = ActionChains(self.driver).move_to_element(symbolInfoContainer)
                    hover.perform()

                    marketDepthWindow = self.getChildElement(symbolInfoContainer, By.XPATH, './div/span/span[3]')
                    marketDepthWindow.click()

                # Data Feeder ============================================================================
                ohlcContainer = self.getChildElement(symbolInfoContainer, By.CSS_SELECTOR, "div[class='ohlc']")

                prevClose = self.getChildElement(ohlcContainer, By.XPATH, './div[2]/div[2]/span').text
                open = self.getChildElement(ohlcContainer, By.XPATH, './div[1]/div[1]/span').text
                high = self.getChildElement(ohlcContainer, By.XPATH, './div[1]/div[2]/span').text
                low = self.getChildElement(ohlcContainer, By.XPATH, './div[2]/div[1]/span').text
                ltp = self.getChildElement(symbolInfoContainer, By.CSS_SELECTOR, "span[class='last-price']").text
                symbol = self.getChildElement(symbolInfoContainer, By.CSS_SELECTOR, "span[class='nice-name']").text

                if symbol != 'NIFTY BANK' and symbol != 'NIFTY 50':
                    bidPrice = self.getChildElement(symbolInfoContainer, By.XPATH, './div/div[2]/div[1]/div[1]/table[1]/tbody/tr[1]/td[1]').text
                    bidQty = self.getChildElement(symbolInfoContainer, By.XPATH, './div/div[2]/div[1]/div[1]/table[1]/tbody/tr[1]/td[3]').text
                    askPrice = self.getChildElement(symbolInfoContainer, By.XPATH, './div/div[2]/div[1]/div[1]/table[2]/tbody/tr[1]/td[1]').text
                    askQty = self.getChildElement(symbolInfoContainer, By.XPATH, './div/div[2]/div[1]/div[1]/table[2]/tbody/tr[1]/td[3]').text
                else:
                    bidPrice = '0.0'
                    bidQty = '0'
                    askPrice = '0.0'
                    askQty = '0'

                symbolL.append(symbol)
                prevCloseL.append(prevClose)
                openL.append(open)
                highL.append(high)
                lowL.append(low)
                closeL.append(ltp)
                bidPriceL.append(bidPrice)
                bidQtyL.append(bidQty)
                askPriceL.append(askPrice)
                askQtyL.append(askQty)


                # Filter OHL Strategy =====================================================
                # if open == high or open == low:
                    # print(symbol + '\t' + datetime.now().strftime("%r"))

            dict = {'PrevClose': prevCloseL, 'Open': openL, 'High': highL, 'Low': lowL, 'Close': closeL,
                    'BidPrice': bidPriceL, 'BidQty': bidQtyL, 'AskPrice': askPriceL, 'AskQty': askQtyL}
            self.nifty50Quotes = pd.DataFrame(dict, index=symbolL)
            # print(self.nifty50Quotes)
            # print(nifty50Quotes.at['SBIN', 'Open'])

            '''
            writer = ExcelWriter('PythonExport.xlsx')
            nifty50Quotes.to_excel(writer, 'Sheet1')
            writer.save()
            '''

        except TimeoutException:
            print("Watchlist Task -> Timeout occurred")

if __name__ == "__main__":
    obj = AutoTrade()
    obj.automation()
