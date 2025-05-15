import os
import time
import json
import re
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, InvalidCookieDomainException


class SNSYCrawler:
    def __init__(self):
        self.login_url = "https://authserver.snsy.edu.cn/authserver/login?service=http%3A%2F%2Fjwgl.snsy.edu.cn%3A8080%2Feams%2Flogin.action"
        self.course_table_url = "http://jwgl.snsy.edu.cn:8080/eams/courseTableForStd.action"
        self.home_url = "http://jwgl.snsy.edu.cn:8080/eams/home.action"
        self.cookie_file = "cookies.json"
        self.driver = None

    def init_driver(self):
        """初始化浏览器驱动"""
        chrome_options = Options()
        # chrome_options.add_argument('--headless')
        # chrome_options.add_argument('--disable-gpu')
        # chrome_options.add_argument('--window-size=1920,1080')

        service = Service()
        driver = webdriver.Chrome(service=service, options=chrome_options)
        self.driver = driver
        return driver

    def convert_csv_to_excel(self, csv_path, excel_path):
        """将CSV文件转换为Excel文件"""
        try:
            df = pd.read_csv(csv_path)
            df.to_excel(excel_path, index=False)
            print(f"成功将 {csv_path} 转换为 {excel_path}")
            return True
        except Exception as e:
            print(f"转换文件时出错: {e}")
            return False

    def _is_cookie_valid(self, cookies):
        """检查cookie是否有效"""
        if not cookies:
            return False
        required_fields = ['name', 'value', 'domain']
        return all(all(field in c for field in required_fields) for c in cookies)

    def load_cookies(self):
        """加载并验证cookies"""
        try:
            if not os.path.exists(self.cookie_file):
                return False

            with open(self.cookie_file, "r", encoding='utf-8') as f:
                cookies = json.load(f)

            if not self._is_cookie_valid(cookies):
                return False

            # 必须先访问域名才能设置cookie
            self.driver.get("http://jwgl.snsy.edu.cn:8080/")
            self.driver.delete_all_cookies()

            for cookie in cookies:
                try:
                    # 修正domain字段（去掉开头的点）
                    domain = cookie['domain'].lstrip('.')
                    self.driver.add_cookie({
                        'name': cookie['name'],
                        'value': cookie['value'],
                        'domain': domain,
                        'path': cookie.get('path', '/'),
                        'secure': cookie.get('secure', False)
                    })
                except InvalidCookieDomainException:
                    # 如果域名不匹配，尝试强制设置为当前域名
                    self.driver.add_cookie({
                        'name': cookie['name'],
                        'value': cookie['value'],
                        'domain': 'jwgl.snsy.edu.cn',
                        'path': '/'
                    })
                except Exception as e:
                    print(f"跳过无效cookie: {str(e)}")
                    continue

            # 直接跳转课表页面验证
            self.driver.get(self.course_table_url)
            if "login.action" not in self.driver.current_url:
                try:
                    WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "table.gridtable"))
                    )
                    print("Cookies验证成功，已自动跳转课表页面")
                    return True
                except:
                    pass

            # 自动登录失败，保存当前的cookies
            current_cookies = self.driver.get_cookies()
            if current_cookies:
                with open(self.cookie_file, "w", encoding='utf-8') as f:
                    json.dump(current_cookies, f, ensure_ascii=False, indent=2)
                print(f"自动登录失败，已将新的Cookies保存到{self.cookie_file}")

            return False
        except Exception as e:
            print(f"加载Cookies失败: {str(e)}")
            return False

    def manual_login(self):
        """手动登录"""
        print("请手动登录系统...")
        self.driver.get(self.login_url)
        try:
            WebDriverWait(self.driver, 300).until(
                lambda d: "home.action" in d.current_url or
                          "courseTableForStd.action" in d.current_url
            )
            print("登录成功!")
            cookies = self.driver.get_cookies()
            with open(self.cookie_file, "w", encoding='utf-8') as f:
                json.dump(cookies, f, ensure_ascii=False, indent=2)
            print(f"Cookies已保存到{self.cookie_file}")
            return True
        except TimeoutException:
            raise Exception("登录超时")

    def navigate_to_course_table(self):
        """导航到课表页面并选择正确的学期"""
        try:
            print("正在导航到课表页面...")
            self.driver.get(self.course_table_url)

            # 等待页面加载
            time.sleep(3)

            # 查找页面上的所有按钮，找到导航到具体课表页面的按钮
            buttons = self.driver.find_elements(By.TAG_NAME, "button")
            links = self.driver.find_elements(By.TAG_NAME, "a")

            # 尝试查找包含"课表"或"我的课表"的链接或按钮
            table_link = None
            for element in links + buttons:
                text = element.text.strip()
                if "课表" in text or "我的课表" in text or "课程表" in text:
                    print(f"找到课表链接: {text}")
                    table_link = element
                    break

            # 如果找到了链接，点击它
            if table_link:
                table_link.click()
                time.sleep(2)
                print("已点击课表链接")

            # 检查是否存在"查询"按钮，通常用于提交表单
            query_buttons = self.driver.find_elements(By.XPATH, "//button[contains(text(), '查询')]")
            if not query_buttons:
                query_buttons = self.driver.find_elements(By.XPATH,
                                                          "//input[@type='submit' or @type='button'][@value='查询']")

            if query_buttons:
                print(f"找到查询按钮: {query_buttons[0].text or query_buttons[0].get_attribute('value')}")
                query_buttons[0].click()
                time.sleep(2)
                print("已点击查询按钮")

            # 保存当前页面，以便检查
            with open("course_table_page.html", "w", encoding="utf-8") as f:
                f.write(self.driver.page_source)
            print("已保存课表页面到course_table_page.html")

            return True

        except Exception as e:
            print(f"导航到课表页面时发生错误: {e}")
            return False

    def extract_course_table(self):
        """提取课程表数据，通用重构版"""
        try:
            print("正在提取课程表...")
            html_content = self.driver.page_source
            soup = BeautifulSoup(html_content, 'html.parser')

            # 查找课表表格
            course_table = soup.find('table', id='manualArrangeCourseTable')
            if not course_table:
                print("无法找到课表表格")
                return []

            # 获取表头信息 - 星期
            headers = course_table.find_all('th')
            weekdays = []
            for i in range(1, len(headers)):  # 跳过第一个表头（节次/周次）
                day_text = headers[i].get_text(strip=True)
                weekdays.append(day_text)

            # 获取所有行，并建立行到时间段的映射
            rows = course_table.find('tbody').find_all('tr')
            time_slots = []
            for row in rows:
                time_cell = row.find('td')
                if time_cell:
                    time_slot = time_cell.get_text(strip=True)
                    time_slots.append(time_slot)

            # 创建一个二维矩阵表示课表，初始值为None
            num_rows = len(time_slots)
            num_cols = len(weekdays)
            course_matrix = [[None for _ in range(num_cols)] for _ in range(num_rows)]

            # 首先通过DOM结构构建正确的位置映射
            # 获取表中每个单元格的位置信息
            for row_idx, row in enumerate(rows):
                cells = row.find_all('td')
                col_offset = 0  # 用于追踪列偏移

                for cell_idx, cell in enumerate(cells):
                    if cell_idx == 0:  # 跳过第一列（时间段列）
                        continue

                    actual_col = cell_idx - 1 + col_offset
                    if actual_col >= num_cols:
                        continue  # 防止超出界限

                    # 如果当前位置已经有内容，说明被之前的单元格的rowspan或colspan覆盖了
                    while actual_col < num_cols and course_matrix[row_idx][actual_col] is not None:
                        actual_col += 1
                        col_offset += 1

                    if actual_col >= num_cols:
                        continue  # 再次检查是否超出界限

                    # 处理跨行跨列
                    rowspan = int(cell.get('rowspan', 1))
                    colspan = int(cell.get('colspan', 1))

                    # 如果是包含课程的单元格
                    if 'infoTitle' in cell.get('class', []):
                        # 提取课程信息
                        course_text = cell.text.strip()
                        weekday = weekdays[actual_col]

                        # 对于跨行的课程，为每个跨越的时间段生成记录
                        for r in range(rowspan):
                            if row_idx + r >= num_rows:
                                continue  # 防止超出行界限

                            time_slot = time_slots[row_idx + r]
                            course_info = self.parse_course_info_improved(cell, course_text, time_slot, weekday)

                            # 将课程信息填入矩阵
                            for c in range(colspan):
                                if actual_col + c < num_cols:
                                    course_matrix[row_idx + r][actual_col + c] = course_info

                    # 无论是否有课程，都需要更新占用情况
                    for r in range(rowspan):
                        if row_idx + r >= num_rows:
                            continue
                        for c in range(colspan):
                            if actual_col + c < num_cols:
                                # 如果没有课程信息，用空列表占位
                                if course_matrix[row_idx + r][actual_col + c] is None:
                                    course_matrix[row_idx + r][actual_col + c] = []

                    # 更新列偏移
                    col_offset += colspan - 1

            # 收集所有课程信息
            all_courses = []
            for row in range(num_rows):
                for col in range(num_cols):
                    if course_matrix[row][col] and course_matrix[row][col]:  # 非空且有内容
                        all_courses.extend(course_matrix[row][col])

            print(f"成功提取 {len(all_courses)} 条课程信息")

            # 保存课程数据
            self.save_course_data_improved(all_courses)

            return all_courses

        except Exception as e:
            print(f"提取课程表时出错: {e}")
            import traceback
            traceback.print_exc()
            return []

    def parse_course_info_improved(self, cell, course_text, time_slot, weekday):
        """改进版解析单个课程单元格的信息，支持拆分不同周次的课程"""
        course_infos = []

        # 从title属性解析完整信息
        title = cell.get('title', '')
        if title:
            # 分号分隔不同的课程信息块
            segments = [seg.strip() for seg in title.split(';') if seg.strip()]

            # 按段解析：通常每对段落包含一个完整的课程信息
            i = 0
            while i < len(segments):
                course_info = {
                    "课程名称": "",
                    "教师": "",
                    "周次": "",
                    "教室": "",
                    "星期": weekday,
                    "节次": time_slot,
                    "单双周": ""
                }

                # 解析课程名和教师
                if i < len(segments):
                    course_teacher_part = segments[i]
                    course_match = re.match(r'(.+?)\((.+?)\) \((.+?)\)', course_teacher_part)
                    if course_match:
                        course_info["课程名称"] = course_match.group(1).strip()
                        # course_code = course_match.group(2).strip()  # 课程代码，暂不保存
                        course_info["教师"] = course_match.group(3).strip()
                    else:
                        # 尝试其他匹配模式
                        name_match = re.match(r'(.+?)\(', course_teacher_part)
                        if name_match:
                            course_info["课程名称"] = name_match.group(1).strip()

                        teacher_match = re.search(r'\) \((.+?)\)$', course_teacher_part)
                        if teacher_match:
                            course_info["教师"] = teacher_match.group(1).strip()
                    i += 1

                # 解析周次和教室
                if i < len(segments) and segments[i].startswith('(') and ')' in segments[i]:
                    week_room_part = segments[i].strip('()')
                    parts = week_room_part.split(',')
                    if len(parts) >= 1:
                        course_info["周次"] = parts[0].strip()
                        # 判断单双周
                        if "单" in course_info["周次"]:
                            course_info["单双周"] = "单周"
                        elif "双" in course_info["周次"]:
                            course_info["单双周"] = "双周"

                        if len(parts) >= 2:
                            course_info["教室"] = parts[1].strip()
                    i += 1
                else:
                    i += 1

                # 如果解析到了有效的课程信息，添加到结果列表
                if course_info["课程名称"]:
                    course_infos.append(course_info)

        # 如果从title属性没有解析到信息，尝试从文本内容解析
        if not course_infos and course_text:
            lines = [line.strip() for line in course_text.split('\n') if line.strip()]

            course_name = ""
            teacher_name = ""

            # 尝试从第一行提取课程名和教师
            if lines and '(' in lines[0] and ')' in lines[0]:
                first_line = lines[0]
                name_match = re.match(r'(.+?)\(', first_line)
                if name_match:
                    course_name = name_match.group(1).strip()

                teacher_match = re.search(r'\) \((.+?)\)$', first_line)
                if teacher_match:
                    teacher_name = teacher_match.group(1).strip()

            # 尝试从后续行提取周次和教室信息
            for i in range(1, len(lines)):
                if i < len(lines) and lines[i].startswith('(') and lines[i].endswith(')'):
                    course_info = {
                        "课程名称": course_name,
                        "教师": teacher_name,
                        "周次": "",
                        "教室": "",
                        "星期": weekday,
                        "节次": time_slot,
                        "单双周": ""
                    }

                    week_room_part = lines[i][1:-1]  # 去掉括号
                    parts = week_room_part.split(',')
                    if len(parts) >= 1:
                        course_info["周次"] = parts[0].strip()
                        # 判断单双周
                        if "单" in course_info["周次"]:
                            course_info["单双周"] = "单周"
                        elif "双" in course_info["周次"]:
                            course_info["单双周"] = "双周"

                        if len(parts) >= 2:
                            course_info["教室"] = parts[1].strip()

                    if course_info["课程名称"]:
                        course_infos.append(course_info)

        return course_infos

    def save_course_data_improved(self, course_data):
        """将课程数据保存为CSV和Excel文件，改进版"""
        try:
            # 创建DataFrame
            df = pd.DataFrame(course_data)

            # 保存为CSV
            csv_path = "course_schedule.csv"
            df.to_csv(csv_path, index=False, encoding='utf-8-sig')
            print(f"课程表已保存为CSV文件: {csv_path}")

            # 转换为Excel
            excel_path = "原始信息.xlsx"
            self.convert_csv_to_excel(csv_path, excel_path)

            return True
        except Exception as e:
            print(f"保存课程数据时出错: {e}")
            return False

    def run(self):
        """运行爬虫"""
        try:
            self.init_driver()

            # 尝试加载cookies
            auto_login_success = self.load_cookies()
            if not auto_login_success:
                # 需要手动登录
                if not self.manual_login():
                    return

            # 导航到课表页面
            if not self.navigate_to_course_table():
                return

            # 提取课表数据
            course_data = self.extract_course_table()

            return course_data

        except Exception as e:
            print(f"运行爬虫时出错: {e}")
            import traceback
            traceback.print_exc()
        finally:
            if self.driver:
                self.driver.quit()


if __name__ == "__main__":
    crawler = SNSYCrawler()
    crawler.run()
