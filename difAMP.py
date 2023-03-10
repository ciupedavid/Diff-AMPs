import unittest
import os
import shutil

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from mpmath import *

import time

import json


# This script downloads the Nimble data and LTSpice data from Beta-tools-analog
# Note: This script uses PYAUTOGUI which controls the mouse for the drag and drop action.
# While this script is running, please do not move the mouse

class TestNimble(unittest.TestCase):

    def setUp(self):
        # driver instance
        options = Options()
        options.headless = True
        options.add_argument("--headless=new")
        self.driver = webdriver.Chrome(options=options)
        with open(r'Devicee.json') as d:
            self.nimbleData = json.load(d)['Nimble'][0]

    def test_export(self):
        driver = self.driver
        driver.set_window_size(1920, 1080)
        driver.get('https://beta-tools.analog.com/noise/#session=SG1ma-6s0EO8c-Wjj2rrGw&step=h16s8AE8RrmSd-LiGDtX9A')
        # cookies accept
        time.sleep(8)
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "body.ember-application:nth-child(2) div.consent-dialog:nth-child(1) div.modal.fade.in.show div.modal-dialog div.modal-content div.modal-body div.short-description > a.btn.btn-success:nth-child(2)"))).click()
        time.sleep(1)
        # sensor settings
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "body.ember-application:nth-child(2) div.tab-content:nth-child(2) div.signal-chain-row:nth-child(2) div.signal-chain-row-item:nth-child(1) div.voltage-diff-src-sensor-container.sc-sensor-container:nth-child(2) > div.sensor-svg.voltage-diff-src-sensor-svg"))).click()
        time.sleep(1)
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#resistance-1-input"))).send_keys(Keys.CONTROL + "a")
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#resistance-1-input"))).send_keys(Keys.DELETE)
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#resistance-1-input"))).send_keys(self.nimbleData['resistance_input'])
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#capacitance-1-input"))).send_keys(Keys.CONTROL + "a")
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#capacitance-1-input"))).send_keys(Keys.DELETE)
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#capacitance-1-input"))).send_keys(self.nimbleData['capacitance_input'])
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "body.ember-application.modal-open:nth-child(2) div.adi-modal:nth-child(4) div.modal.fade.in.show:nth-child(1) div.modal-dialog div.modal-content form.modal-footer div.button-row > button.btn.btn-primary:nth-child(1)"))).click()
        time.sleep(1)
        # amplifier settings
        driver.execute_script("document.querySelector('#signal-chain-drop-area #circuit-content[title=\"Amplifier\"]').click()")
        time.sleep(1)
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#amp-gain-input"))).click()
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#amp-gain-input"))).send_keys(Keys.CONTROL + "a")
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#amp-gain-input"))).send_keys(Keys.DELETE)
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#amp-gain-input"))).send_keys(self.nimbleData['gain'])
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#amp-gain-input"))).send_keys(Keys.ENTER)
        time.sleep(1)
        position = value_to_position(self.nimbleData['scale_selector'])
        self.scrollToValue(position, driver)
        self.scrollToCValue(self.nimbleData['c1_scale'], driver)
        time.sleep(1)
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#tspan2988-4-54-5"))).click()
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#filter-0"))).send_keys(self.nimbleData['device'])
        time.sleep(1)
        driver.execute_script("document.querySelector('div.slick-cell.l0.r0.frozen').click();")
        driver.execute_script("document.querySelector('.modal-footer button.btn-primary').click();")
        time.sleep(1)
        driver.execute_script("document.querySelector('.modal-footer button.btn-primary').click();")
        time.sleep(1)
        driver.execute_script("document.querySelector('#signal-chain-drop-area #circuit-content[title=\"Filter\"]').click()")
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#filter-inputs-type-tab-button"))).click()
        time.sleep(1)
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#hp-diff-wiring-button"))).click()
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR,"body.ember-application.modal-open:nth-child(2) div.adi-modal.modal-fills-window:nth-child(5) div.modal.fade.in.show:nth-child(1) div.modal-dialog div.modal-content div.modal-body div.configure-filter.configure-signal-chain-item div.top-area section.config-section div.adi-sub-tab-container div.sub-tab-content-container div.sub-tab-content:nth-child(2) div.sub-tab-input-config div.config-input-row:nth-child(2) div.adi-radio-group > div.adi-radio:nth-child(1)"))).click()
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#fp-input"))).send_keys(Keys.CONTROL + "a")
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#fp-input"))).send_keys(Keys.DELETE)
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#fp-input"))).send_keys(self.nimbleData['filter_frequency'])
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR,"body.ember-application.modal-open:nth-child(2) div.adi-modal.modal-fills-window:nth-child(5) div.modal.fade.in.show:nth-child(1) div.modal-dialog div.modal-content div.modal-body div.configure-filter.configure-signal-chain-item div.top-area section.config-section div.adi-sub-tab-container div.sub-tab-content-container button.tab-button-area.enabled.next:nth-child(3) > div.tab-button.enabled.next"))).click()
        time.sleep(3)

        self.scrollToRCValue(self.nimbleData['rc_scale'], driver)
        time.sleep(3)
        self.scrollToRC3Value(self.nimbleData['rc3_scale'], driver)
        time.sleep(3)
        driver.execute_script("document.querySelector(\"form[class='modal-footer'] button:nth-child(1)\").click()")
        time.sleep(1)
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#next-steps-tab"))).click()
        time.sleep(1)
        device = self.nimbleData['device']
        downloads_path = self.nimbleData['downloads_path']
        gain = self.nimbleData['gain']
        current_date = self.nimbleData['current_date']
        l = driver.current_url
        device_url = device + 'URL G' + gain + '.txt'
        with open(device_url, 'w') as f:
            f.write(l)
        print(l)
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR,"body.ember-application:nth-child(2) div.tab-content:nth-child(2) div.download-area div.download-individual-buttons div.download-button-row:nth-child(1) button.btn.btn-primary:nth-child(1) > span:nth-child(1)"))).click()
        time.sleep(4)
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR,"body.ember-application:nth-child(2) div.tab-content:nth-child(2) div:nth-child(1) div.download-area div.download-individual-buttons div.download-button-row:nth-child(1) > button.btn.btn-primary:nth-child(2)"))).click()
        time.sleep(5)
        # project_path = os.getcwd()

        project_path = self.nimbleData['project_location']

        if not os.path.exists(project_path + '\\' + device):
            os.makedirs(project_path + '\\' + device)
        dir_list = os.listdir()
        print(dir_list)
        print(project_path + device)

        ltspice_download_path = downloads_path + 'LTspice ' + current_date + '.zip'
        shutil.move(ltspice_download_path, project_path + '/' + device)
        print(ltspice_download_path, project_path + '/' + device)
        nimble_download_path = downloads_path + 'Raw Data Export - ' + current_date + '.zip'
        shutil.move(nimble_download_path, project_path + '/' + device)
        device_download_path = device + 'URL G' + gain + '.txt'
        shutil.move(device_download_path, project_path + '/' + device)
        # shutil.move(os.path.join(device_download_path, device + 'URL G' + gain + '.txt'), os.path.join(project_path + '\\' + device, device + 'URL G' + gain + '.txt'))
        time.sleep(1)

        downloaded_nimble_path = project_path + '\\' + device + '\\' + 'Raw Data Export - ' + current_date + '.zip'
        new_nimble_name = project_path + '\\' + device + '\\' + 'Nimble - ' + device + ' G' + gain + '.zip'
        downloaded_ltspice_path = project_path + '\\' + device + '\\' + 'LTspice ' + current_date + '.zip'
        new_ltspice_name = project_path + '\\' + device + '\\' + 'LTspice - ' + device + ' G' + gain + '.zip'
        os.rename(downloaded_nimble_path, new_nimble_name)
        os.rename(downloaded_ltspice_path, new_ltspice_name)

        time.sleep(2)


    @staticmethod
    def scrollToValue(value: int, driver):
        driver.execute_script(f"document.querySelector('#rscale-slider').value = {value}; document.querySelector('#rscale-slider').dispatchEvent(new Event('input'));")

    @staticmethod
    def scrollToCValue(value: int, driver):
        driver.execute_script(
            f"document.querySelector('#c1-slider').value = {value}; document.querySelector('#c1-slider').dispatchEvent(new Event('input'));")

    @staticmethod
    def scrollToRCValue(value: int, driver):
        driver.execute_script(f"document.querySelector('#rc-r1-slider').value = {value}; document.querySelector('#rc-r1-slider').dispatchEvent(new Event('input'));")

    @staticmethod
    def scrollToRC3Value(value: int, driver):
        driver.execute_script(
            f"document.querySelector('#rc-r3-slider').value = {value}; document.querySelector('#rc-r3-slider').dispatchEvent(new Event('input'));")

    time.sleep(2)

    def tearDown(self):
        self.driver.quit()


if __name__ == '__main__':
    unittest.main()


def value_to_position(value):
    # html/js values
    minpos = 1
    maxpos = 10000
    # slider values
    minimum_slider_value = 10
    maximum_slider_value = 10000000
    # logarith function
    minval = log(minimum_slider_value)
    maxval = log(maximum_slider_value)
    # scale & postion equations
    scale = (maxval - minval) / (maxpos - minpos)
    rposition = minpos + (log(value) - minval) / scale
    return rposition
