# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.
import os
import time
from datetime import datetime
from time import sleep

import threading

import xlsxwriter
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys


class LOCATORS:
    ID = "id"
    CLASS_NAME = "class_name"
    XPATH = "xpath"
    LINK_TEXT = "link_text"


def delete_existing_value_and_enter_new_value(ele, value):
    for _ in range(4):
        ele.send_keys(Keys.BACKSPACE)
        sleep(0.1)
    ele.send_keys(value)
    ele.send_keys(Keys.ENTER)


def read_strategy_output_performance(output):
    final_output = []
    for loop in range(len(output[0])):
        profit = sum(float(inner_list[loop][2]) for inner_list in output)
        trades = sum(float(inner_list[loop][3]) for inner_list in output)
        profit_percent = sum(
            float(inner_list[loop][4].split()[0]) for inner_list in output
        ) / len(output)

        profit_factor = sum(float(inner_list[loop][5]) for inner_list in output) / len(
            output
        )

        draw_dwn = max(float(inner_list[loop][6]) for inner_list in output)
        average_trade = sum(float(inner_list[loop][7]) for inner_list in output) / len(
            output
        )

        average_bars = sum(float(inner_list[loop][8]) for inner_list in output) / len(
            output
        )

        final_output.append(
            [
                output[0][loop][0],
                output[0][loop][1],
                profit,
                trades,
                profit_percent,
                profit_factor,
                draw_dwn,
                average_trade,
                average_bars,
            ]
        )
    return final_output


class Script:
    def __init__(
            self,
            pyramiding_start,
            pyramiding_end,
            chart,
            deep_history,
            step_size_start,
            step_size_end,
            time_frame_start,
            time_frame_end,
            step_jump,
            commission,
            slippage,
    ):

        self.driver = webdriver.Chrome(
            executable_path=f"{os.getcwd()}/resources/chromedriver"
        )
        # self.driver=webdriver.Firefox(executable_path=f"{os.getcwd()}/resources/geckodriver")
        # self.driver = webdriver.Safari()
        self.weblink = "https://in.tradingview.com/"
        self.chart = chart
        self.strategy = "Renko Overlay - pune3tghai"
        self.pyramiding_start = pyramiding_start
        self.pyramiding_end = pyramiding_end
        self.deep_history = deep_history
        self.step_size_start = step_size_start
        self.step_size_end = step_size_end
        self.time_frame_start = time_frame_start
        self.time_frame_end = time_frame_end
        self.step_jump = step_jump
        self.commission = commission
        self.slippage = slippage

    def open_weblink(self):
        self.driver.maximize_window()
        self.driver.get(self.weblink)
        self.driver.maximize_window()
        print(f"{self.pyramiding_start}-{self.pyramiding_end} open weblink")

    def get_element(
            self, element_locator, locator_identifier=LOCATORS.XPATH, ele_name=""
    ):
        retry = 1
        element = None
        while retry < 10:
            try:
                if locator_identifier == LOCATORS.XPATH:
                    element = self.driver.find_element_by_xpath(element_locator)
                elif locator_identifier == LOCATORS.ID:
                    element = self.driver.find_element_by_id(element_locator)
                elif locator_identifier == LOCATORS.CLASS_NAME:
                    element = self.driver.find_element_by_class_name(element_locator)
                elif locator_identifier == LOCATORS.LINK_TEXT:
                    element = self.driver.find_element_by_link_text(element_locator)
                return element
            except Exception as e:
                try:
                    self.driver.find_element_by_xpath("//span[text()='Connect']").click()
                except:
                    pass
                if ele_name:
                    print(
                        f"{self.pyramiding_start}-{self.pyramiding_end} attempt: {retry} accessing element: {ele_name}"
                    )
                else:
                    print(
                        f"{self.pyramiding_start}-{self.pyramiding_end} attempt: {retry} accessing element: {element_locator}"
                    )
                retry += 1
                sleep(1)

    def wait_for_date_to_refresh(self):
        refresh_data_time = 1
        while refresh_data_time < 150:
            try:
                if self.deep_history:
                    self.driver.find_element(
                        value="//div[@class='backtesting deep-history']//div[@role='progressbar']",
                        by=By.XPATH,
                    )
                else:
                    self.driver.find_element_by_class_name(
                        "reports-content opacity-transition fade"
                    )
                refresh_data_time += 1
                sleep(1)
                print(
                    f"{self.pyramiding_start}-{self.pyramiding_end} data is refreshing"
                )
            except:
                sleep(3)
                print(
                    f"{self.pyramiding_start}-{self.pyramiding_end} data has been refreshed"
                )
                break

    def click_element(
            self, element_locator, locator_identifier=LOCATORS.XPATH, ele_name=""
    ):
        if isinstance(element_locator, str):
            print(
                f"{self.pyramiding_start}-{self.pyramiding_end} clicking {ele_name or element_locator}"
            )
            self.click_element(self.get_element(
                element_locator,
                locator_identifier=locator_identifier,
                ele_name=ele_name,
            ), ele_name=ele_name)
            print(
                f"{self.pyramiding_start}-{self.pyramiding_end} clicked {ele_name or element_locator}"
            )
        else:
            retry = 1
            while retry < 10:
                try:
                    print(
                        f"{self.pyramiding_start}-{self.pyramiding_end} clicking {ele_name or element_locator}"
                    )
                    element_locator.click()
                    print(
                        f"{self.pyramiding_start}-{self.pyramiding_end} clicked {ele_name or element_locator}"
                    )
                    break
                except:
                    try:
                        self.driver.find_element_by_xpath("//span[text()='Connect']").click()
                    except:
                        pass
                    sleep(1)
                    retry += 1
            else:
                print(
                    f"{self.pyramiding_start}-{self.pyramiding_end} unable to click: {element_locator}"
                )

    def send_keys(self, element_locator, text, locator_identifier=LOCATORS.XPATH):
        self.get_element(
            element_locator, locator_identifier=locator_identifier
        ).send_keys(text)

    def perform_actions(self, keys):
        actions = ActionChains(self.driver)
        actions.send_keys(keys)
        sleep(2)
        print(f"{self.pyramiding_start}-{self.pyramiding_end} Performing Actions!")
        actions.perform()
        sleep(3)

    def delete_cache(self):
        self.driver.execute_script(
            "window.open('')"
        )  # Create a separate tab than the main one
        self.driver.switch_to.window(
            self.driver.window_handles[-1]
        )  # Switch window to the second tab
        self.driver.get(
            "chrome://settings/clearBrowserData"
        )  # Open your chrome settings.
        self.perform_actions(
            Keys.TAB * 4 + Keys.DOWN * 5 + Keys.TAB * 5 + Keys.ENTER,
        )  # Tab to the time select and key down to say "All Time" then go to the Confirm button and press Enter
        sleep(2)
        self.driver.close()  # Close that window
        self.driver.switch_to.window(
            self.driver.window_handles[0]
        )  # Switch Selenium controls to the original tab to continue normal functionality.

    def login(self):
        # start here

        sleep(1)
        sign_in_icon_ele = self.get_element(
            "//div[@class='tv-header__area tv-header__area--user']/button[1]",
            LOCATORS.XPATH,
        )
        self.driver.execute_script("arguments[0].click();", sign_in_icon_ele)
        print(f"{self.pyramiding_start}-{self.pyramiding_end} clicked sign in icon")
        sleep(1)

        sign_in_text = self.get_element("//span[text()='Sign in']")
        self.driver.execute_script("arguments[0].click();", sign_in_text)
        print(f"{self.pyramiding_start}-{self.pyramiding_end} clicked sign in text")
        sleep(1)

        self.click_element("//span[text()='Email']", ele_name="Email")
        self.get_element(
            "//input[starts-with(@id,'email-signin__user-name')]"
        ).send_keys("maheshk00100@gmail.com")
        print(f"{self.pyramiding_start}-{self.pyramiding_end} entered username")
        self.get_element(
            "//input[starts-with(@id,'email-signin__password')]"
        ).send_keys("MAHkok@100")
        print(f"{self.pyramiding_start}-{self.pyramiding_end} entered password")
        self.click_element(
            "tv-button__loader", LOCATORS.CLASS_NAME, ele_name="Final Sign in"
        )
        sleep(3)

    def close_driver(self):
        self.driver.quit()

    def select_chart(self):
        self.click_element("//button[@aria-label='Search']")
        self.send_keys("//input[@type='search']", self.chart)
        print(f"{self.pyramiding_start}-{self.pyramiding_end} Entered: {self.chart}")
        sleep(1)
        try:
            self.driver.find_element_by_xpath("//span[text()='Connect']").click()
        except:
            pass
        self.send_keys("//input[@type='search']", Keys.RETURN)
        print(
            f"{self.pyramiding_start}-{self.pyramiding_end} Hit Entered on select chart"
        )
        sleep(3)

    def get_strategy_performance(self):
        result = []
        for index in range(1, 8):
            value = self.get_element(
                f"//div[@class='container-b1pZpka9']/div[{index}]/div[2]/div[1]",
                ele_name="report data",
            ).text
            if "−" in value:
                value = value.replace("−", "-")
            if "INR" in value:
                result.append("".join(value.split()[:-1]))
            else:
                result.append(value)
        return result

    def one_time_setup(self, pyramiding, time_frame):
        mouse = ActionChains(self.driver)
        mouse.send_keys(time_frame).send_keys(Keys.ENTER)
        sleep(0.5)
        mouse.perform()
        sleep(1)
        while True:
            deep_hisotry_ele = self.get_element(
                '//input[@type="checkbox"][@role="switch"]',
                ele_name="deep history element",
            )
            if deep_hisotry_ele.get_attribute("aria-checked") == "false":
                self.click_element(
                    '//input[@type="checkbox"][@role="switch"]',
                    ele_name="deep history element",
                )
                sleep(1)
            else:
                break
        sleep(1)

        self.click_element("//span[text()='Accept all']", ele_name="Accept All Cookies")
        # row = 0
        todays_date = datetime.now().date()
        # Enter Date
        for loop in range(1, 3):
            date_ele = f"(//div[@class='container-HGdxdr5y']//input[@class='input-oiYdY6I4 with-end-slot-oiYdY6I4'])[{loop}]"
            for _ in range(10):
                self.send_keys(
                    date_ele,
                    Keys.BACKSPACE,
                )
                sleep(0.1)
            self.send_keys(
                date_ele,
                str(todays_date.replace(year=todays_date.year - 1))
                if loop == 1
                else str(todays_date),
            )

        # Code for step size, and we rely on every candle hence its not used at the moment
        self.click_strategy_settings_icon()
        # enter pyramiding value
        self.click_element("//div[@data-value='properties']")
        pyramiding_ele = self.get_element(
            "//div[@class='content-mTbR5jYu']/div[8]//input"
        )
        delete_existing_value_and_enter_new_value(pyramiding_ele, pyramiding)

        commission_ele = self.get_element(
            "//div[@class='content-mTbR5jYu']/div[11]//input"
        )
        delete_existing_value_and_enter_new_value(commission_ele, self.commission)

        slippage_ele = self.get_element(
            "//div[@class='content-mTbR5jYu']/div[15]//input"
        )
        delete_existing_value_and_enter_new_value(slippage_ele, self.slippage)

        # click inputs
        self.click_element("//div[@data-value='inputs']")
        sleep(0.2)

        mv_input_ele_index = [9, 11]
        # for setting open mv and close mv
        mv_list = [11, 12]
        for loop, mv in enumerate(mv_input_ele_index):
            mv_input_ele = f"//div[@class='content-mTbR5jYu']/div[{mv}]//input"
            for outer_i in range(4):
                self.get_element(mv_input_ele).send_keys(Keys.BACKSPACE)
                sleep(0.1)
            self.get_element(mv_input_ele).send_keys(mv_list[loop])
            sleep(0.1)

        self.click_element("//button[@name='submit']", ele_name="ok button")
        sleep(1)

    def select_strategy(self):
        self.click_element(
            "header-toolbar-indicators", locator_identifier=LOCATORS.ID, ele_name="fx "
        )
        sleep(1)
        self.send_keys("//input[@placeholder='Search']", self.strategy)
        print(
            f"{self.pyramiding_start}-{self.pyramiding_end} Enter: {self.strategy} strategy to search for"
        )
        sleep(1)
        self.click_element("//div[starts-with(@class,'main')]")
        print(f"{self.pyramiding_start}-{self.pyramiding_end} Clicked strategy")
        sleep(1)
        self.click_element("//span[@data-name='close']")
        print(
            f"{self.pyramiding_start}-{self.pyramiding_end} select strategy screen closed"
        )

    def enter_step_size(self, step_size):
        brick_size_ele = self.get_element("(//input[@inputmode='numeric'])[1]")
        brick_size_ele.click()
        delete_existing_value_and_enter_new_value(brick_size_ele, step_size)
        self.click_element("//button[@name='submit']", ele_name="ok button")
        sleep(1)

    def click_strategy_settings_icon(self):
        try:
            self.click_element(
                "//div[@class='backtesting deep-history']/div[1]/div[1]/div[1]/div[2]/button[1]/span[1]",
                ele_name="Tester Area strategy settings icon",
            )
            sleep(1)
        except Exception as e:
            pass

    def click_generate_report_and_get_strategy_results(self):
        self.click_element(
            "//span[text()='Generate report']",
            LOCATORS.XPATH,
            ele_name="Generate Report",
        )
        sleep(3)
        self.wait_for_date_to_refresh()
        sleep(1)
        return self.get_strategy_performance()

    def evaluate_best_results(self):
        datetime_format = "%Y_%d-%m_%H_%M_%S_%f"

        time_frame_start = self.time_frame_start
        while time_frame_start < self.time_frame_end:
            self.select_time_frame(time_frame_start)

            with xlsxwriter.Workbook(
                    f"{os.getcwd()}/output/{str(datetime.utcnow().strftime(datetime_format)).replace(':', '')}_{self.chart}-TF-{time_frame_start}-{self.pyramiding_start}-{self.pyramiding_end - 1}.xlsx",
                    {"constant_memory": True, "strings_to_numbers": True},
            ) as workbook:
                for pyramiding in range(self.pyramiding_start, self.pyramiding_end):
                    worksheet = workbook.add_worksheet(f"pyramiding_{pyramiding}")
                    output = []
                    try:
                        step_size = self.step_size_start
                        attempt = 1
                        self.one_time_setup(pyramiding, time_frame_start)
                        _iter = 1
                        while step_size <= self.step_size_end:
                            try:
                                self.click_strategy_settings_icon()
                                self.enter_step_size(step_size)
                                strategy_results = (
                                    self.click_generate_report_and_get_strategy_results()
                                )
                                strategy_results = [
                                    time_frame_start,
                                    step_size,
                                    *strategy_results,
                                ]
                                print(
                                    f"{self.pyramiding_start} - {self.pyramiding_end} Pyramiding: {pyramiding}, Brick Size: {step_size}, strategy details: {strategy_results}"
                                )

                                output.append(strategy_results)
                                step_size += self.step_jump
                                attempt = 0
                                sleep(1)

                                if _iter == 5:
                                    self.delete_cache_and_login()
                                    self.one_time_setup(pyramiding, time_frame_start)
                                    _iter = 1
                                else:
                                    _iter += 1

                            except Exception as e:
                                print(
                                    f"{self.pyramiding_start}-{self.pyramiding_end} exception occurred on strategy settings screen",
                                    e,
                                )
                                if attempt == 5:
                                    step_size += self.step_jump
                                    attempt = 0
                                else:
                                    attempt += 1

                        sorted_output = sorted(
                            output, key=lambda x: float(x[2]), reverse=True
                        )
                        sorted_output.insert(
                            0,
                            [
                                "time_frame",
                                "step_size",
                                "profit",
                                "total_trades",
                                "profit %",
                                "profit factor",
                                "max dradowm",
                                "avg trade",
                                "avg bars",
                            ],
                        )
                        self.write_to_excel_sheet(sorted_output, worksheet)

                    except Exception as e:
                        final_output = read_strategy_output_performance(output)
                        for row, inner_list in enumerate(final_output):
                            try:
                                for col_index, data in enumerate(inner_list):
                                    try:
                                        worksheet.write_row(row, col_index, data)
                                    except:
                                        pass
                            except:
                                pass

                        print(
                            f"{self.pyramiding_start}-{self.pyramiding_end} Exception occurred: e",
                            e,
                        )

                    self.delete_cache_and_login()
            time_frame_start = time_frame_start + 1

    @staticmethod
    def write_to_excel_sheet(sorted_output, worksheet):
        for row, inner_list in enumerate(sorted_output):
            for col_index, data in enumerate(inner_list):
                worksheet.write(row, col_index, data)
        print("data written to excel")

    def select_time_frame(self, time_frame):
        mouse = ActionChains(self.driver)
        mouse.send_keys(time_frame)
        sleep(0.2)
        mouse.send_keys(Keys.ENTER)
        sleep(0.2)
        mouse.perform()
        sleep(0.5)
        print(
            f"{self.pyramiding_start}-{self.pyramiding_end} selected: {time_frame} time frame"
        )

    def delete_cache_and_login(self):
        self.delete_cache()
        self.driver.refresh()
        try:
            self.driver.switch_to.alert.accept()
        except Exception:
            print("Alert not present")
        sleep(2)
        self.login()
        sleep(5)


def open_webpage(driver, weblink):
    driver.get(weblink)


def entire_run(**kwargs):
    if missing_keys := set(kwargs.keys()) - {
        "pyramiding_start",
        "pyramiding_end",
        "step_size_start",
        "step_size_end",
        "chart",
        "deep_history",
        "time_frame_start",
        "time_frame_end",
        "step_jump",
        "commission",
        "slippage",
    }:
        raise Exception(f"missing keys: {missing_keys}")

    pyramiding_start = kwargs["pyramiding_start"]
    pyramiding_end = kwargs["pyramiding_end"]
    step_size_start = kwargs["step_size_start"]
    step_size_end = kwargs["step_size_end"]
    chart = kwargs["chart"]
    deep_history = kwargs["deep_history"]
    time_frame_start = kwargs["time_frame_start"]
    time_frame_end = kwargs["time_frame_end"]
    step_jump = kwargs["step_jump"]
    commission = kwargs["commission"]
    slippage = kwargs["slippage"]

    print(f"pyramiding_start: {pyramiding_start}")
    print(f"pyramiding_end: {pyramiding_end}")
    print(f"step_size_start: {step_size_start}")
    print(f"step_size_end: {step_size_end}")
    print()
    script = Script(
        pyramiding_start,
        pyramiding_end,
        chart=chart,
        deep_history=deep_history,
        step_size_start=step_size_start,
        step_size_end=step_size_end,
        time_frame_start=time_frame_start,
        time_frame_end=time_frame_end,
        step_jump=step_jump,
        commission=commission,
        slippage=slippage,
    )
    script.open_weblink()
    try:
        script.login()
        script.select_chart()
        script.evaluate_best_results()
    except Exception as e:
        print(f"{pyramiding_start}-{pyramiding_end} Exception occurred: ", e)
        script.close_driver()
    script.close_driver()


if __name__ == "__main__":

    kwargs = {
        "step_jump": 1,
        "deep_history": True,
        "time_frame_start": 2,
        "time_frame_end": 3,
    }

    nifty_kwargs = {
        "step_size_start": 6,
        "step_size_end": 80,
        "chart": "NIFTY1!",
        "commission": 160,
        "slippage": 125,
    }

    bank_nifty_kwargs = {
        "step_size_start": 16,
        "step_size_end": 100,
        "chart": "BANKNIFTY1!",
        "commission": 160,
        "slippage": 250,
    }
    no_of_threads = 5
    pyramiding_start = 1
    pyramiding_end = 100

    pyramiding_segments = list(
        range(
            pyramiding_start,
            pyramiding_end + 1,
            ((pyramiding_end + 1) - pyramiding_start) // no_of_threads,
        )
    )
    thread_list = []
    for thread in range(no_of_threads):
        new_kwargs = {
            **kwargs,
            **bank_nifty_kwargs,
            "pyramiding_start": pyramiding_segments[thread],
            "pyramiding_end": pyramiding_segments[thread + 1]
            if thread != no_of_threads - 1
            else pyramiding_end,
        }

        t = threading.Thread(
            target=entire_run,
            kwargs=new_kwargs,
        )
        new_kwargs = {}
        thread_list.append(t)
        new_kwargs = {}

    [_thread.start() for _thread in thread_list]
    [_thread.join() for _thread in thread_list]
