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
                        value="//div[@class='backtesting deep-history']//div[@role='progressbar']", by=By.XPATH
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

    def initiate_script(self):
        # create directory range_filter
        # create yearwise folder

        for sampling_period in range(50):
            workbook = xlsxwriter.Workbook(f"output/{sampling_period}.xlsx")
            # create excelsheet and set its name as per sampling period

            range_multiplier_numbers = [x * 0.1 for x in range(1, 31)]
            for range_multiplier in range_multiplier_numbers:
                worksheet = workbook.add_worksheet(f"sheet_{range_multiplier}"[:9])
                # create sheet based on constant ranging from 0.1 to 20 and set its name as per constant value
                row = 0
                column = 0
                for time_frame in range(1, 120):
                    for data in range(10):
                        worksheet.write(row, column, data)
                        column += 1
                    column = 0
                    row += 1
                    # set time_frame
                    # extract data and dump it to sheet

            workbook.close()

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
                f"//div[@class='container-b1pZpka9']/div[{index}]/div[2]/div[1]", ele_name="report data"
            ).text
            if '−' in value:
                value = value.replace('−', '-')
            if "INR" in value:
                result.append("".join(value.split()[:-1]))
            else:
                result.append(value)
        return result

        # final_data = []
        # for data_index in range(1, 8):
        #     ele_xpath = f"//div[@class='report-data']/div[1]/div[{data_index}]/div[2]/div[1]"
        #     if data_index != 5:
        #         data = self.driver.find_element_by_xpath(ele_xpath).text
        #     else:
        #         data = self.driver.find_element_by_xpath(f"{ele_xpath}/span").text
        #     final_data.append(data)
        #
        # return final_data

    def delete_existing_value_and_enter_new_value(self, ele, value):
        for _ in range(4):
            ele.send_keys(Keys.BACKSPACE)
            sleep(0.1)
        ele.send_keys(value)
        ele.send_keys(Keys.ENTER)

    def one_time_setup(self, pyramiding, time_frame):
        # if (
        #     time_frame
        #     not in self.get_element(
        #         "//div[@id='header-toolbar-intervals']/div/div/div"
        #     ).text
        # ):
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
        self.click_element(
            "//div[@class='backtesting deep-history']/div[1]/div[1]/div[1]/div[2]/button[1]/span[1]",
            ele_name="Tester Area strategy settings icon",
        )
        sleep(1)
        # enter pyramiding value
        self.click_element("//div[@data-value='properties']")
        pyramiding_ele = self.get_element(
            "//div[@class='content-mTbR5jYu']/div[8]//input"
        )
        self.delete_existing_value_and_enter_new_value(pyramiding_ele, pyramiding)

        commission_ele = self.get_element(
            "//div[@class='content-mTbR5jYu']/div[11]//input"
        )
        self.delete_existing_value_and_enter_new_value(commission_ele, self.commission)

        slippage_ele = self.get_element(
            "//div[@class='content-mTbR5jYu']/div[15]//input"
        )
        self.delete_existing_value_and_enter_new_value(slippage_ele, self.slippage)

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

    def read_strategy_output_performance(self, output):
        final_output = []
        for loop in range(len(output[0])):
            profit = sum(float(inner_list[loop][2]) for inner_list in output)
            trades = sum(float(inner_list[loop][3]) for inner_list in output)
            profit_percent = sum(
                float(inner_list[loop][4].split()[0]) for inner_list in output
            ) / len(output)

            profit_factor = sum(
                float(inner_list[loop][5]) for inner_list in output
            ) / len(output)

            draw_dwn = max(float(inner_list[loop][6]) for inner_list in output)
            average_trade = sum(
                float(inner_list[loop][7]) for inner_list in output
            ) / len(output)

            average_bars = sum(
                float(inner_list[loop][8]) for inner_list in output
            ) / len(output)

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

    def enter_step_size(self, step_size):
        brick_size_ele = self.get_element("(//input[@inputmode='numeric'])[1]")
        brick_size_ele.click()
        self.delete_existing_value_and_enter_new_value(brick_size_ele, step_size)
        self.click_element("//button[@name='submit']", ele_name="ok button")
        sleep(1)

    def evaluate_best_results(self):
        datetime_format = "%Y_%d-%m_%H_%M_%S_%f"
        if not self.deep_history:
            mouse = ActionChains(self.driver)
            with xlsxwriter.Workbook(
                f"{os.getcwd()}/output/{datetime.utcnow().strftime(datetime_format)}_{self.chart}-{self.pyramiding_start}-{self.pyramiding_end}.xlsx",
                {"constant_memory": True, "strings_to_numbers": True},
            ) as workbook:
                for pyramiding in range(self.pyramiding_start, self.pyramiding_end):
                    worksheet = workbook.add_worksheet(f"pyramiding_{pyramiding}")
                    # row = 0
                    time_frame = 5
                    while time_frame < 6:
                        date_range = [
                            [2021, 5, 1, 2021, 12, 1],
                            [2021, 12, 1, 2022, 10, 15],
                        ]
                        output = [[] for _ in range(len(date_range))]
                        try:
                            # if time_frame < 10:
                            #     mouse.move_to_element(
                            #         self.get_element("(//div[@class='chart-gui-wrapper'])[1]")
                            #     ).send_keys(1).send_keys(Keys.BACKSPACE).send_keys(
                            #         time_frame
                            #     ).send_keys(
                            #         Keys.ENTER
                            #     ).perform()
                            # elif 10 <= time_frame <= 99:
                            #     mouse.move_to_element(
                            #         self.get_element("(//div[@class='chart-gui-wrapper'])[1]")
                            #     ).send_keys(1).send_keys(Keys.BACKSPACE).send_keys(
                            #         Keys.BACKSPACE
                            #     ).send_keys(
                            #         time_frame
                            #     ).send_keys(
                            #         Keys.ENTER
                            #     ).perform()
                            # else:
                            #     mouse.move_to_element(
                            #         self.get_element("(//div[@class='chart-gui-wrapper'])[1]")
                            #     ).send_keys(1).send_keys(Keys.BACKSPACE).send_keys(
                            #         Keys.BACKSPACE
                            #     ).send_keys(
                            #         Keys.BACKSPACE
                            #     ).send_keys(
                            #         time_frame
                            #     ).send_keys(
                            #         Keys.ENTER
                            #     ).perform()
                            # sleep(10)
                            # strategy_on_chart_ele = self.get_element(
                            #     f"//div[text()='{self.strategy}']"
                            # )

                            # self.click_element(strategy_on_chart_ele)

                            # chart_gui_ele = self.get_element("(//div[@class='chart-gui-wrapper'])[1]")
                            # mouse.move_to_element(chart_gui_ele).click().perform()
                            # sleep(0.2)
                            # self.send_keys("(//div[@class='chart-gui-wrapper'])[1]", time_frame)
                            # sleep(0.3)
                            # self.send_keys("(//div[@class='chart-gui-wrapper'])[1]", Keys.RETURN)
                            # sleep(10)
                            # print(f"Time Frame: {time_frame} selected")

                            # self.click_element(strategy_on_chart_ele, ele_name="strategy_on_chart")
                            # sleep(0.2)
                            # self.click_element("//div[@data-name='legend-settings-action']", ele_name="settings icon")

                            self.click_element(
                                # "//div[@class='backtesting-head-wrapper']/div[2]/div",
                                "//div[@class='backtesting deep-history']/div[1]/div[1]/div[1]/div[2]/button[1]/span[1]",
                                ele_name="Tester Area strategy settings icon",
                            )
                            # mouse.move_to_element(self.get_element("//div[@data-name='legend-settings-action']")).click()
                            # sleep(0.2)
                            # mouse.double_click(self.get_element("//div[@data-name='legend-settings-action']")).perform()
                            # print("clicked strategy setting icons")
                            sleep(1)

                            # # self.click_element("icon-button js-backtesting-open-format-dialog apply-common-tooltip", LOCATORS.CLASS_NAME, ele_name="strategy setting icon")
                            # mini_container_ele = self.get_element("report-minichart-container", LOCATORS.CLASS_NAME)
                            # mouse.move_to_element(mini_container_ele).context_click(mini_container_ele).perform()
                            # sleep(0.2)
                            # self.driver.execute_script("arguments[0].click();", self.get_element("//span[text()='Strategy Properties…']"))
                            # print("clicked strategy properties")
                            # self.click_element("//span[text()='Strategy Properties…']", ele_name="Strategy Properties")
                            # sleep(0.5)
                            # self.click_element(f"//div[text()='{self.strategy}']", ele_name="strategy on chart")
                            # sleep(1)
                            #
                            # print("opening strategy's setting screen")
                            # mouse.double_click(self.get_element(f"//div[text()='{self.strategy}']"))
                            # print("opened strategy's setting screen")
                            # # print("clicking strategy's setting icon ")
                            # # self.driver.execute_script("arguments[0].click();", self.get_element("//div[@data-name='legend-settings-action']"))
                            # # print("clicked strategy's setting icon")
                            # self.click_element("//div[@data-name='legend-settings-action']", ele_name="strategy's setting icon")

                            for inner_list_index, date in enumerate(date_range):
                                step_size = 5
                                attempt = 1

                                # enter pyramiding value
                                self.click_element("//div[@data-value='properties']")
                                pyramiding_ele = self.get_element(
                                    "//div[@class='cell-ByXdMGQj'][4]//input"
                                )
                                for _ in range(4):
                                    pyramiding_ele.send_keys(Keys.BACKSPACE)
                                    sleep(0.1)
                                pyramiding_ele.send_keys(pyramiding)
                                pyramiding_ele.send_keys(Keys.ENTER)

                                # click inputs
                                self.click_element("//div[@data-value='inputs']")
                                sleep(0.2)

                                date_input_eles = self.driver.find_elements_by_xpath(
                                    "//input[@class='input-uGWFLwEy with-end-slot-uGWFLwEy']"
                                )

                                # for setting open mv and close mv
                                mv_list = [11, 12]
                                for mv_index, mv in enumerate(mv_list):
                                    for outer_i in range(4):
                                        date_input_eles[mv_index + 3].send_keys(
                                            Keys.BACKSPACE
                                        )
                                        sleep(0.1)
                                    date_input_eles[mv_index + 3].send_keys(mv)
                                    sleep(0.1)

                                # set date range
                                for index, value in enumerate(date):
                                    for _ in range(4):
                                        date_input_eles[index + 5].send_keys(
                                            Keys.BACKSPACE
                                        )
                                        sleep(0.1)
                                    date_input_eles[index + 5].send_keys(value)
                                    date_input_eles[index + 5].send_keys(Keys.ENTER)
                                    sleep(0.1)

                                while step_size <= 60:
                                    try:
                                        brick_size_ele = self.get_element(
                                            "(//input[@inputmode='numeric'])[1]"
                                        )
                                        brick_size_ele.click()
                                        brick_size_ele.send_keys(Keys.BACKSPACE)
                                        sleep(0.1)
                                        brick_size_ele.send_keys(Keys.BACKSPACE)
                                        sleep(0.1)
                                        brick_size_ele.send_keys(Keys.BACKSPACE)
                                        sleep(0.1)
                                        self.get_element(
                                            "(//input[@inputmode='numeric'])[1]"
                                        ).send_keys(step_size)
                                        sleep(0.1)
                                        self.get_element(
                                            "(//input[@inputmode='numeric'])[1]"
                                        ).send_keys(Keys.RETURN)
                                        sleep(1)
                                        self.wait_for_date_to_refresh()
                                        sleep(1)
                                        strategy_output_details = [
                                            time_frame,
                                            step_size,
                                            *self.get_strategy_performance(),
                                        ]
                                        # for col_index, d in enumerate(final_data):
                                        #     worksheet.write(row, col_index, d)
                                        # row += 1
                                        if output[inner_list_index]:
                                            max_time_to_wait = 1
                                            while max_time_to_wait <= 15:
                                                if (
                                                    strategy_output_details[2]
                                                    == output[inner_list_index][-1][2]
                                                ):
                                                    self.wait_for_date_to_refresh()
                                                    strategy_output_details = [
                                                        time_frame,
                                                        step_size,
                                                        *self.get_strategy_performance(),
                                                    ]
                                                    max_time_to_wait += 1
                                                    sleep(1)
                                                else:
                                                    break

                                        print(f"{strategy_output_details}")
                                        output[inner_list_index].append(
                                            strategy_output_details
                                        )
                                        step_size += 1
                                        attempt = 0
                                    except Exception as e:
                                        print(
                                            "exception occurred on strategy settings screen",
                                            e,
                                        )
                                        if attempt == 5:
                                            step_size += 1
                                            attempt = 0
                                        else:
                                            attempt += 1

                            final_output = []
                            for loop in range(len(output[0])):
                                profit = sum(
                                    [
                                        float(inner_list[loop][2])
                                        for inner_list in output
                                    ]
                                )
                                trades = sum(
                                    [
                                        float(inner_list[loop][3])
                                        for inner_list in output
                                    ]
                                )
                                profit_percent = sum(
                                    [
                                        float(inner_list[loop][4].split()[0])
                                        for inner_list in output
                                    ]
                                ) / len(output)
                                profit_factor = sum(
                                    [
                                        float(inner_list[loop][5])
                                        for inner_list in output
                                    ]
                                ) / len(output)
                                draw_dwn = max(
                                    [
                                        float(inner_list[loop][6])
                                        for inner_list in output
                                    ]
                                )
                                average_trade = sum(
                                    [
                                        float(inner_list[loop][7])
                                        for inner_list in output
                                    ]
                                ) / len(output)
                                average_bars = sum(
                                    [
                                        float(inner_list[loop][8])
                                        for inner_list in output
                                    ]
                                ) / len(output)
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

                            for row, inner_list in enumerate(final_output):
                                for col_index, data in enumerate(inner_list):
                                    worksheet.write(row, col_index, data)

                            self.click_element("//span[@data-name='close']")
                            print(
                                f"{self.pyramiding_start}-{self.pyramiding_end} edit strategy screen closed"
                            )
                            sleep(0.2)
                            time_frame += 1
                        except Exception as e:

                            final_output = []
                            for loop in range(len(output[0])):
                                profit = sum(
                                    [
                                        float(inner_list[loop][2])
                                        for inner_list in output
                                    ]
                                )
                                trades = sum(
                                    [
                                        float(inner_list[loop][3])
                                        for inner_list in output
                                    ]
                                )
                                profit_percent = sum(
                                    [
                                        float(inner_list[loop][4].split()[0])
                                        for inner_list in output
                                    ]
                                ) / len(output)
                                profit_factor = sum(
                                    [
                                        float(inner_list[loop][5])
                                        for inner_list in output
                                    ]
                                ) / len(output)
                                draw_dwn = max(
                                    [
                                        float(inner_list[loop][6])
                                        for inner_list in output
                                    ]
                                )
                                average_trade = sum(
                                    [
                                        float(inner_list[loop][7])
                                        for inner_list in output
                                    ]
                                ) / len(output)
                                average_bars = sum(
                                    [
                                        float(inner_list[loop][8])
                                        for inner_list in output
                                    ]
                                ) / len(output)
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
                            time_frame += 1
        else:
            time_frame_start = self.time_frame_start
            while time_frame_start < self.time_frame_end:
                mouse = ActionChains(self.driver)
                mouse.send_keys(time_frame_start).send_keys(Keys.ENTER)
                sleep(0.1)
                mouse.perform()
                sleep(1)
                print(
                    f"{self.pyramiding_start}-{self.pyramiding_end} selected: {time_frame_start} time frame"
                )
                with xlsxwriter.Workbook(
                    f"{os.getcwd()}/output/{str(datetime.utcnow().strftime(datetime_format)).replace(':', '')}_{self.chart}-TF-{time_frame_start}-{self.pyramiding_start}-{self.pyramiding_end-1}.xlsx",
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
                                    self.click_element(
                                        # "//div[@class='backtesting-head-wrapper']/div[2]/div",
                                        "//div[@class='backtesting deep-history']/div[1]/div[1]/div[1]/div[2]/button[1]/span[1]",
                                        ele_name="Tester Area strategy settings icon",
                                    )
                                    sleep(1)

                                    # brick_size_ele = self.get_element(
                                    #     "(//input[@inputmode='numeric'])[1]"
                                    # )
                                    # brick_size_ele.click()
                                    # brick_size_ele.send_keys(Keys.BACKSPACE)
                                    # sleep(0.1)
                                    # brick_size_ele.send_keys(Keys.BACKSPACE)
                                    # sleep(0.1)
                                    # brick_size_ele.send_keys(Keys.BACKSPACE)
                                    # sleep(0.1)
                                    # self.get_element(
                                    #     "(//input[@inputmode='numeric'])[1]"
                                    # ).send_keys(step_size)
                                    # sleep(0.1)
                                    # self.click_element(
                                    #     "//button[@name='submit']", ele_name="ok button"
                                    # )
                                    # sleep(1)

                                    self.enter_step_size(step_size)

                                    self.click_element(
                                        "//span[text()='Generate report']",
                                        LOCATORS.XPATH,
                                        ele_name="Generate Report",
                                    )
                                    sleep(3)
                                    self.wait_for_date_to_refresh()
                                    sleep(1)
                                    try:
                                        strategy_output_details = [
                                            time_frame_start,
                                            step_size,
                                            *self.get_strategy_performance(),
                                        ]
                                        print(
                                            f"{self.pyramiding_start} - {self.pyramiding_end} Pyramiding: {pyramiding}, Brick Size: {step_size}, strategy details: {strategy_output_details}"
                                        )
                                        output.append(strategy_output_details)
                                        step_size += self.step_jump
                                        attempt = 0
                                        sleep(1)

                                        if _iter == 5:
                                            self.delete_cache()
                                            self.driver.refresh()
                                            try:
                                                self.driver.switch_to.alert.accept()
                                            except Exception:
                                                print("Alert not present")
                                            sleep(2)
                                            self.login()
                                            sleep(5)
                                            self.one_time_setup(
                                                pyramiding, time_frame_start
                                            )
                                            _iter = 1
                                        else:
                                            _iter += 1
                                    except Exception as e:
                                        # chrome runs out of memory when trying to access the strategy report
                                        # hence refresh the page and run from the same loop
                                        print(
                                            f"{self.pyramiding_start}-{self.pyramiding_end} Error: {e}"
                                        )
                                        print(
                                            f"{self.pyramiding_start}-{self.pyramiding_end} chrome out of memory"
                                        )
                                        self.driver.refresh()
                                        self.one_time_setup(
                                            pyramiding, time_frame_start
                                        )
                                        continue

                                except Exception as e:
                                    print(
                                        f"{self.pyramiding_start}-{self.pyramiding_end} exception occurred on strategy settings screen",
                                        e,
                                    )
                                    try:
                                        self.driver.find_element_by_xpath("//span[text()='Connect']").click()
                                    except:
                                        pass
                                    if attempt == 5:
                                        step_size += self.step_jump
                                        attempt = 0
                                    else:
                                        attempt += 1

                            print("output to write to excel: ", output)
                            sorted_output = sorted(
                                output, key=lambda x: float(x[2]), reverse=True
                            )
                            print("sorted output to write to excel: ", sorted_output)
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
                            for row, inner_list in enumerate(sorted_output):
                                for col_index, data in enumerate(inner_list):
                                    worksheet.write(row, col_index, data)
                            print("data written to excel")

                        except Exception as e:
                            final_output = self.read_strategy_output_performance(output)
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

                        self.delete_cache()
                        self.driver.refresh()
                        try:
                            self.driver.switch_to.alert.accept()
                        except Exception as e:
                            print("Alert not present")
                        sleep(2)
                        self.login()
                        sleep(5)
                time_frame_start = time_frame_start + 1


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
            **nifty_kwargs,
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
