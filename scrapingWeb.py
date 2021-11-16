from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, SessionNotCreatedException, WebDriverException

from selenium.webdriver.support.ui import WebDriverWait

from urllib.parse import urlparse

import configparser

config = configparser.ConfigParser()
config.read('config\config.ini')
driverChrome = config['DRIVERS']['CHROME']

class ScrapingWeb:

    def __init__(self, version, url, user, password ):
        self.version = version
        self.url = url
        self.user = user
        self.password = password


    def valid_url(self):
        try:
            result = urlparse(self.url)
            return all([result.scheme, result.netloc])
        except ValueError:
            return False

    def conecction_nw_70(self):
        try:
            self.driver.get(self.url)
            self.driver.find_element_by_xpath("//input[@id='logonuidfield']").send_keys(self.user)
            self.driver.find_element_by_xpath("//input[@id='logonpassfield']").send_keys(self.password)
            self.driver.find_element_by_xpath("//input[@name='uidPasswordLogon']").click()
            error = "Connection OK"

            try:
                element = self.driver.find_element_by_id('welcome_message')
                error = "Connection OK"
            except NoSuchElementException as err:
                error = "No ingreso al portal"

                try:
                    element = self.driver.find_element_by_class_name('urMsgBarErr')
                    errMsg = "//span[@class='urTxtStd']"

                    element = self.driver.find_element_by_xpath(errMsg)

                    error = element.get_attribute('innerHTML')

                except NoSuchElementException as err:
                    error = "No encontro mensaje de error en login"

        except NoSuchElementException as err:
            error = "No se tuvo acceso al portal"
        except WebDriverException as err:
            error = "No se tuvo acceso, Time Out URL: " + self.url +" "+ str(err)

        return error


    def conecction_NW_75(self):
        try:
            self.driver.get(self.url)
            self.driver.find_element_by_xpath("//input[@id='logonuidfield']").send_keys(self.user)
            self.driver.find_element_by_xpath("//input[@id='logonpassfield']").send_keys(self.password)
            self.driver.find_element_by_xpath("//input[@name='uidPasswordLogon']").click()
            error = "Connection OK"

            try:
                element = self.driver.find_element_by_id('CEPJ.IDPView.Welcome')
                error = "Connection OK"
            except NoSuchElementException as err:
                error = "No ingreso al portal"

                try:
                    element = self.driver.find_element_by_class_name('urMsgBarErr')
                    errMsg = "//span[@class='urTxtMsg']"

                    element = self.driver.find_element_by_xpath(errMsg)

                    error = element.get_attribute('innerHTML')

                except NoSuchElementException as err:
                    error = "No encontro mensaje de error en login"

        except NoSuchElementException as err:
            error = "No se tuvo acceso al portal"
        except WebDriverException as err:
            error = "No se tuvo acceso, Time Out URL: " + self.url +" "+ str(err)

        return error

    def conecction_BO(self):

        try:
            self.driver.get(self.url)

            self.driver.switch_to_default_content()

            iframe = self.driver.find_element_by_xpath("//iframe")
            self.driver.switch_to.frame(iframe)

            # self.driver.find_element_by_xpath("//input[@id='_id2: logon:CMS']").send_keys(self.CMS)
            self.driver.find_element_by_xpath("//input[@id='_id2:logon:USERNAME']").send_keys(self.user)
            self.driver.find_element_by_xpath("//input[@id='_id2:logon:PASSWORD']").send_keys(self.password)
            self.driver.find_element_by_xpath("//input[@id='_id2:logon:logonButton']").click()
            error = "Connection OK"

            try:
                element = self.driver.find_element_by_id('logoleft')
                error = "Connection OK"
            except NoSuchElementException as err:
                error = "No ingreso al portal"

                try:
                    element = self.driver.find_element_by_class_name('logonError')
                    errMsg = "//div[@class='logonError']"

                    element = self.driver.find_element_by_xpath(errMsg)

                    error = element.get_attribute('innerHTML')

                except NoSuchElementException as err:
                    error = "No encontro mensaje de error en login"

        except NoSuchElementException as err:
            error = "No se tuvo acceso al portal"
        except WebDriverException as err:
            error = "No se tuvo acceso, Time Out URL: " + self.url +" "+ str(err)

        return error

    def conecction_FIORI(self):

        try:
            self.driver.get(self.url)

            self.driver.find_element_by_xpath("//input[@id='USERNAME_FIELD-inner']").send_keys(self.user)
            self.driver.find_element_by_xpath("//input[@id='PASSWORD_FIELD-inner']").send_keys(self.password)
            self.driver.find_element_by_xpath("//button[@id='LOGIN_LINK']").click()
            error = "Connection OK"

            try:
                element = self.driver.find_element_by_id('logoleft')
                error = "Connection OK"
            except NoSuchElementException as err:
                error = "No ingreso al portal"

                try:
                    element = self.driver.find_element_by_xpath("//div[@id='LOGIN_ERROR_BLOCK']")
                    errMsg = "//label[@id='LOGIN_LBL_ERROR']"

                    element = self.driver.find_element_by_xpath(errMsg)

                    error = element.get_attribute('innerHTML')

                except NoSuchElementException as err:
                    error = "No encontro mensaje de error en login"

        except NoSuchElementException as err:
            error = "No se tuvo acceso al portal"
        except WebDriverException as err:
            error = "No se tuvo acceso, Time Out URL: " + self.url +" "+ str(err)

        return error

    def test_connection(self):
        if self.valid_url():
            options = Options()
            options.add_argument('--ignore-ssl-errors=yes')
            options.add_argument('--allow-insecure-localhost')
            options.add_argument('--allow-running-insecure-content')
            options.add_argument('--ignore-certificate-errors')


            ##
            #ChromeOptions
            #options = new
            #ChromeOptions()
            #chrome_options.add_argument('--allow-insecure-localhost')
            #DesiredCapabilities caps = DesiredCapabilities.chrome()
            #caps.setCapability(ChromeOptions.CAPABILITY, options)
            #caps.setCapability("acceptInsecureCerts", true)
            #WebDriverdriver = new ChromeDriver(caps)
            ###

            #driver = webdriver.Chrome()
            try:
                self.driver = webdriver.Chrome(executable_path=driverChrome, chrome_options=options)

                if (self.version == "NW 7.0"):
                    error = self.conecction_nw_70()
                elif (self.version == "NW 7.5"):
                    error = self.conecction_NW_75()
                elif (self.version == "BO"):
                    error = self.conecction_BO()
                elif (self.version == "FIORI"):
                    error = self.conecction_FIORI()
                else:
                    error = "Sistema no soportado " + self.version

                self.driver.close()
            except SessionNotCreatedException as err:
                error = "Existe un problema con la versión del driver de Chrome. " + str(err)
        else:
            error = "URL incorrecta"
        return error