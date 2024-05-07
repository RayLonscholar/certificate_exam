from flet import *
from flet_route import Params, Basket
from assets.pages.navigation import rail_item
import openpyxl
import random
import time

class Test(UserControl):
    def __init__(self, page):
        super().__init__()

        # 重新選擇題庫
        def restart_clicked(e):
            self.exam_settings.visible = True # 顯示第一區塊
            self.get_results.visible = False # 隱藏提交考卷按鈕
            self.open_question_navigation.visible = False # 隱藏題目列表按鈕
            self.question_navigation.visible = False # 隱藏答題資訊欄
            self.question_information_main.visible = False # 隱藏第二區塊
            self.check_result_main.visible = False # 隱藏第三區塊
            write_list() # 刷新工作表選擇
            page.update()

        self.test_restart = TextButton(
            content = Container(
                content=Column(
                    controls = [
                        Text(value="重新選擇題庫", size=20),
                    ],
                    alignment=MainAxisAlignment.CENTER,
                    spacing=5,
                ),
            ),
            on_click=restart_clicked)
        
        # 提交考卷
        def get_results_clicked(e):
            # 寫入此測驗的題目和答題答案
            self.history_excel = openpyxl.load_workbook("./assets/data/history_data.xlsx") # 讀取excel檔
            self.history_wb = self.history_excel["答題記錄"]
            # 算分數
            # print(self.test_topic_data[self.question_num[0]-1])
            right = 0 # 對的題數
            for index, ans_content in enumerate(self.all_my_answers):
                if ''.join(ans_content) == self.test_topic_data[self.question_num[index]-1][2].strip( ):
                    right += 1
            # 修改第三區塊的元件內容
            self.check_result_main.controls[0].value = f"測驗結果： 答對{right}題"
            # 清空元件
            self.check_result.controls = []
            # 列出對和錯的題目
            for index, question_num in enumerate(self.question_num):
                data = self.test_topic_data[question_num-1]
                order = data[0].split("!n")
                # 題目
                if ''.join(self.all_my_answers[index]) == data[2].strip( ): # 判斷是否答錯
                    self.check_result.controls.append(
                        Row(
                            controls = [
                                Icon(name=icons.CIRCLE_OUTLINED, color="green"),
                                Text("{}.({})：{}".format(index+1, data[2], order[0][4:]), size = 20, weight = "bold"),
                            ],
                        )
                    )
                else:
                    self.check_result.controls.append(
                        Row(
                            controls = [
                                Icon(name=icons.CLOSE, color="red"),
                                Text("{}.({})：{}".format(index+1, data[2], order[0][4:]), size = 20, weight = "bold"),
                            ],
                        )
                    )
                # 題目圖片
                if data[1] != None:
                    self.check_result.controls.append(
                        Image(src=f"{data[1]}")
                    )
                for _ in order[1:]:
                    self.check_result.controls.append(
                    Text(f"{_}", size = 20, weight = "bold")
                    )
                
                self.check_result.controls.append(
                    Text("", size = 20, weight = "bold")
                )

                # 選項
                # for index, _ in enumerate(content[3:]):
                #     option = ['A', 'B', 'C', 'D']
                #     if _ != None:
                #         if _[0] in option: # 判斷是否有圖片
                #             self.topic_information.controls.append(
                #                 Text(f"{_}", size = 20, weight = "bold")
                #             )
                #         else:
                #             self.topic_information.controls.append(
                #                 Row(
                #                     controls = [
                #                         Text(f"{option[index]}：", size = 20, weight = "bold"),
                #                         Image(src=f"{_}")
                #                     ],
                #                 )
                #             )
                
            # self.check_result.controls.append()

            self.exam_settings.visible = False # 隱藏第一區塊
            self.get_results.visible = False # 隱藏提交考卷按鈕
            self.open_question_navigation.visible = False # 隱藏題目列表按鈕
            self.question_navigation.visible = False # 隱藏答題資訊欄
            self.question_information_main.visible = False # 隱藏第二區塊
            self.check_result_main.visible = True # 顯示第三區塊
            page.update()
            # self.question_num.insert(0, )
            # self.history_wb["A1"].value = self.history_wb["A1"].value.insert() # 每一次答題的題目紀錄
            # self.history_wb["A2"].value = self.history_wb["A2"].value.insert(self.all_my_answers) # 每一次答題的答案紀錄
            # print(now_history_position)


        self.get_results = TextButton(
            visible = False,
            content = Container(
                content=Column(
                    controls = [
                        Text(value="提交考卷", size=20),
                    ],
                    alignment=MainAxisAlignment.CENTER,
                    spacing=5,
                ),
            ),
            on_click=get_results_clicked)
        
        # 答題資訊欄
        self.question_navigation_num = GridView( # 題號按鈕區
            expand = True,
            max_extent= 85,
            controls = [],
        )
        self.question_navigation = Column(
            visible = False,
            # expand = True,
            width = 200,
            controls = [
                Text("題號：", size = 20, weight = "bold"),
                self.question_navigation_num,
            ],
        )
        def open_question_navigation_click(e): # 開/關答題資訊欄
            self.question_navigation.visible = not self.question_navigation.visible
            page.update()
        self.open_question_navigation = Row(
            visible = False,
            expand = True,
            alignment = MainAxisAlignment.END,
            controls = [
                TextButton(
                    content = Container(
                        content=Column(
                            controls = [
                                Text(value="題目列表", size=20),
                            ],
                            alignment=MainAxisAlignment.CENTER,
                            spacing=5,
                        ),
                    ),
                    on_click=open_question_navigation_click
                )
            ],
        )

        # 開始抽題並開始作答題目
        def test_submit_clicked(e):
            # 提示對話
            def warning_info(text:str):
                info = AlertDialog(
                    title=Text(text)
                )
                page.dialog = info
                info.open = True
                page.update()
            
            # 生成題目
            def create_question(question_index):
                # 取得當前題目
                self.question_data = self.test_topic_data[self.question_num[question_index]-1]
                # 清空元件
                self.question_information.controls = []
                self.question_information_up_down_button.controls = []
                # 題目
                question_order = self.question_data[0].split("!n")
                self.question_information.controls.append(
                    Text("{}.{}".format(question_index+1, question_order[0]), size = 20, weight = "bold")
                )
                # 題目圖片
                if self.question_data[1] != None:
                    self.question_information.controls.append(
                        Image(src=f"{self.question_data[1]}")
                    )
                for _ in question_order[1:]:
                    self.question_information.controls.append(
                    Text(f"{_}", size = 20, weight = "bold")
                    )

                self.question_information.controls.append(
                    Text("", size = 20, weight = "bold")
                )

                # 判斷答案並推測題目類型
                def to_my_answer(e, mode):# 記錄所選的答案
                    if mode == "單選":
                        _ = ['A', 'B', 'C', 'D']
                        # print(self.ch.controls)
                        for i in self.ch.controls: # 只有一個選項能勾
                            try:
                                # 把不是選中項的取消
                                if i.data == e.control.data:
                                    i.value = True
                                    i.controls[0].value = True
                                else:
                                    i.value = False
                                    i.controls[0].value = False 
                                # 圖片選項的處理(如同上)
                                if i.controls[0]:
                                    if i.controls[0].data == e.control.data:
                                        i.value = True
                                        i.controls[0].value = True # 圖片的選項
                                    else:
                                        i.value = False
                                        i.controls[0].value = False # 圖片的選項
                            except AttributeError: # 沒有圖片的選項
                                # print("AttributeError")
                                pass
                        ans = _[int(e.control.data)]
                        self.all_my_answers[question_index] = [ans] # 記錄此題的答案
                        print("ALL：{}".format(self.all_my_answers))
                    
                    if mode == "複選":
                        _ = ['A', 'B', 'C', 'D']
                        ans = _[int(e.control.data)]
                        if ans not in self.my_answer:
                            self.my_answer.append(ans)
                            # print("add："+ans)
                            # print(self.my_answer)
                        else:
                            self.my_answer.remove(ans)
                            # print("pop："+ans)
                            # print(self.my_answer)
                        self.my_answer.sort()
                        self.all_my_answers[question_index] = self.my_answer # 記錄此題的答案
                        print("ALL：{}".format(self.all_my_answers))
                        
                    if mode == "是非":
                        _ = ['O', 'X']
                        # print(self.ch.controls)
                        for i in self.ch.controls: # 只有一個選項能勾
                            if i.data == e.control.data: # 把不是選中項的取消
                                i.value = True
                            else:
                                i.value = False
                        ans = _[int(e.control.data)]
                        self.all_my_answers[question_index] = [ans] # 記錄此題的答案
                        print("ALL：{}".format(self.all_my_answers))

                    if self.all_my_answers[question_index] != []:
                        # print("change color")
                        self.question_navigation_num.controls[question_index].style = ButtonStyle(bgcolor = colors.TERTIARY_CONTAINER)
                    else:
                        # print("change color = None")
                        self.question_navigation_num.controls[question_index].style = ButtonStyle(bgcolor = None)
                    page.update()
                def choice_question(): # 單選
                    self.ch = Column(
                        controls = [],
                    )
                    for index, _ in enumerate(self.question_data[3:]):
                        # 創建選項
                        option = ['A', 'B', 'C', 'D']
                        if _ != None:
                            if _[0] in option: # 判斷是否有圖片
                                self.ch.controls.append(Checkbox(data = f"{index}", label = f"{_}", on_change = lambda e:to_my_answer(e, "單選")))
                            else:
                                self.ch.controls.append(
                                    Row(
                                        controls = [
                                            Checkbox(data = f"{index}", label = f"{option[index]}：", on_change = lambda e:to_my_answer(e, "單選")),
                                            Image(src=f"{_}"),
                                        ],
                                    )
                                )
                        # 取得此題上次所選的答案
                        for i in self.all_my_answers[question_index]:
                            for j in self.all_my_answers[question_index]:
                                if j == _[0]:
                                    # print("last ans："+j)
                                    self.ch.controls[index].value = True
                    self.question_information.controls.append(self.ch)
                    page.update()
                    # Radio(value = "blue", label = "Blue - Adaptive Radio", adaptive = True, active_color = colors.BLUE),
                def multiple_choice_questions(): # 複選
                    self.ch = Column(
                        controls = [],
                    )
                    for index, _ in enumerate(self.question_data[3:]):
                        # 創建選項
                        option = ['A', 'B', 'C', 'D']
                        if _ != None:
                            if _[0] in option: # 判斷是否有圖片
                                self.ch.controls.append(Checkbox(data = f"{index}", label = f"{_}", on_change = lambda e:to_my_answer(e, "複選")))
                            else:
                                self.ch.controls.append(
                                    Row(
                                        controls = [
                                            Checkbox(data = f"{index}", label = f"{option[index]}：", on_change = lambda e:to_my_answer(e, "複選")),
                                            Image(src=f"{_}")
                                        ],
                                    ) 
                                )
                        # 取得此題上次所選的答案
                        for i in self.all_my_answers[question_index]:
                            # print("i="+i)
                            for j in self.all_my_answers[question_index]:
                                # print("j="+j)
                                if j == _[0]:
                                    # print("last ans："+j)
                                    self.ch.controls[index].value = True
                    self.question_information.controls.append(self.ch)
                    page.update()
                def T_F_question(): # 是非
                    self.ch = Column(
                        controls = [],
                    )
                    for index, _ in enumerate(["O：是", "X：否"]):
                        # 創建選項
                        self.ch.controls.append(Checkbox(data = f"{index}", label = f"{_}", on_change = lambda e:to_my_answer(e, "是非")))
                        # 取得此題上次所選的答案
                        for i in self.all_my_answers[question_index]:
                            for j in self.all_my_answers[question_index]:
                                if j == _[0]:
                                    # print("last ans："+j)
                                    self.ch.controls[index].value = True
                        # 同步已選過的選項使用self.ch.controls的list來改
                        # self.all_my_answers[question_index]
                        
                    self.question_information.controls.append(self.ch)
                    page.update()
                self.my_answer = self.all_my_answers[question_index] # 我選的答案
                self.question_answer = self.question_data[2].strip( ) # 題目正解
                # print(self.question_answer)
                try:
                    if self.question_answer in ['A', 'B', 'C', 'D']: # 單選題
                        choice_question()
                    elif self.question_answer[0] in ['A', 'B', 'C', 'D']: # 複選題
                        multiple_choice_questions()
                    elif self.question_answer[0] in ['O', 'X']: # 是非題
                        T_F_question()
                except TypeError:
                    print("TypeError")
                    warning_info("發生錯誤：請確認題庫資料格式是否正確\n錯誤代碼：TypeError")
                    

                # 上一題、下一題
                def question_information_up_down_button_clicked(e):
                    pass
                    if e.control.data == "上一題":
                        create_question(question_index-1)
                    if e.control.data == "下一題":
                        create_question(question_index+1)
                if question_index > 0:
                    self.question_information_up_down_button.controls.append(
                        TextButton(
                            data = "上一題",
                            content = Container(
                                content = Column(
                                    controls = [
                                        Text(value = f"上一題", size = 20),
                                    ],
                                    alignment = MainAxisAlignment.CENTER,
                                ),
                            ),
                            on_click = lambda e:question_information_up_down_button_clicked(e)
                        )
                    )
                if question_index+1 < len(self.question_num):
                    self.question_information_up_down_button.controls.append(
                        TextButton(
                            data = "下一題",
                            content = Container(
                                content = Column(
                                    controls = [
                                        Text(value = f"下一題", size = 20),
                                    ],
                                    alignment = MainAxisAlignment.CENTER,
                                ),
                            ),
                            on_click = lambda e:question_information_up_down_button_clicked(e)
                        )
                    )
                print(self.question_data)
                print("ALL：{}".format(self.all_my_answers))
                # self.question_information
                # self.test_topic_data
                page.update()

            def question_navigation_num_clikc(e): # 點選答題資訊欄的題號
                # print(e.control.content.content.controls[0].data)
                create_question(int(e.control.content.content.controls[0].data))

            # 判斷輸入的值是否為整數
            try:
                start = round(float(self.test_start.value)) # 開始題號
                end = round(float(self.test_end.value)) # 結束題號
                random_num = int(self.test_random_text.value) # 抽幾題
                if start > 0:
                    if end <= len(self.test_topic_data):
                        # 開始抽題
                        self.question_index = 0 # 題目起始點值
                        self.question_num = random.sample(range(start, end+1), random_num)
                        self.all_my_answers = [[] for _ in range(random_num)] # 我每題作答的答案
                        self.question_navigation_num.controls = [] # 清除答題資訊欄
                        for index, _ in enumerate(self.all_my_answers): # 初始化答題資訊欄
                            self.question_navigation_num.controls.append(
                                TextButton(
                                    content = Container(
                                        content=Column(
                                            controls = [
                                                Text(data = f"{index}", value=f"{index+1}", size=16),
                                            ],
                                            alignment=MainAxisAlignment.CENTER,
                                        ),
                                    ),
                                    on_click = lambda e:question_navigation_num_clikc(e)
                                )
                            )
                        print(self.all_my_answers)
                        create_question(self.question_index)
                        # print(self.question_num)
                        self.exam_settings.visible = False # 隱藏第一區塊
                        self.question_information_main.visible = True # 顯示第二區塊
                        self.get_results.visible = True # 顯示提交考卷按鈕
                        self.open_question_navigation.visible = True # 顯示題目列表按鈕
                        page.update()

                    else:
                        warning_info("請確認輸入的題號最大值是否小於等於總題數") # 提示
                else:
                    warning_info("請確認輸入的題號最小值是否大於0") # 提示

            except ValueError:
                # print("ValueERROR")
                warning_info("請確認輸入的題庫、題號與抽題數是否正確") # 提示

    # 區塊1：選擇題庫
        def load_excel(): # 抓取excel資料
            self.excel = openpyxl.load_workbook("exam_data.xlsx") # 讀取excel檔
            self.choose_test_sheetname.items = []
            page.update()
        def write_list(): # 寫入工作表清單
            load_excel()
            for workbook in self.excel.sheetnames[1:]:
                self.choose_test_sheetname.items.append(PopupMenuItem(text = f"{workbook}", on_click = lambda e:choose_test_sheetname_clicked(e))) 
        def get_values(sheet):
            arr = []
            # 行
            for row in sheet:
                arr2 = []
                # 列
                for column in row:
                    arr2.append(column.value) # 寫入內容
                arr.append(arr2)
            return arr
        def choose_test_sheetname_clicked(e):
            self.choose_test_sheetname_text.value = f"題庫：{e.control.text}"
            self.test_wb = self.excel[e.control.text] # 定位工作表
            self.test_topic_data = get_values(self.test_wb)
            # print(self.test_topic_data)
            self.start_end_range_slider.max = len(self.test_topic_data)
            self.start_end_range_slider.divisions = len(self.test_topic_data)
            self.start_end_range_slider.start_value = 1
            self.start_end_range_slider.end_value = len(self.test_topic_data)
            self.test_start.value = 1
            self.test_end.value = len(self.test_topic_data)
            self.test_start_end.visible = True
            page.update()
        # 確認送出
        self.test_submit = TextButton(
            content = Container(
                content = Column(
                    controls = [
                        Text(value = "確認", size = 20),
                    ],
                    alignment=MainAxisAlignment.CENTER,
                    spacing = 5,
                ),
                padding = padding.all(10),
            ),
            on_click = test_submit_clicked)
        # 選擇題庫
        self.choose_test_sheetname_text = Text("題庫：", size = 20, weight = "bold")
        self.choose_test_sheetname = PopupMenuButton(
            items=[],
            icon = icons.CLEAR_ALL_OUTLINED,
        )
        write_list()
        # 輸入考題數
        def test_start_end_change(e):
            if e.control.data == "開始題號":
                self.start_end_range_slider.start_value = e.control.value
            if e.control.data == "結束題號":
                self.start_end_range_slider.end_value = e.control.value
            if e.control.data == "開始與結束滑塊":
                self.test_start.value = round(float(e.control.start_value))
                self.test_end.value = round(float(e.control.end_value))
            page.update() 
        # 輸入開始題號
        self.test_start = TextField(
            label = "開始題號",
            data = "開始題號",
            width = 100,
            on_change = test_start_end_change
        )
        # 輸入結束題號
        self.test_end = TextField(
            label = "結束題號",
            data = "結束題號",
            width = 100,
            on_change = test_start_end_change
        )
        # 開始與結束題號滑塊
        self.start_end_range_slider = RangeSlider(
            expand = True,
            width = 500,
            data = "開始與結束滑塊",
            min = 1,
            max = 50,
            start_value = 1,
            divisions = 50,
            end_value = 20,
            label="{value}",
            on_change = test_start_end_change
        )
        # 輸入抽題數
        def all_in(e):
            self.test_random_text.value = round(float(self.test_end.value)) - round(float(self.test_start.value)) + 1
            page.update()
        self.test_all_random = TextButton(
            content = Container(
                content = Column(
                    controls = [
                        Text("全部", size = 20),
                    ],
                    alignment = MainAxisAlignment.CENTER,
                ),
            ),
            on_click = lambda e:all_in(e)
        )
        self.test_random_text = TextField(
            label = "抽題數",
            data = "抽題數",
            width = 100,
            on_change = test_start_end_change
        )
        # 開始、結束與抽題
        self.test_start_end = Column(
            visible = False,
            controls = [
                Row(
                    controls = [
                        self.test_start,
                        self.test_end,
                        self.start_end_range_slider,
                    ],
                ),
                Row(
                    controls = [
                        self.test_all_random,
                        self.test_random_text,
                    ],
                )
            ],
        )
        # 主區塊
        self.exam_settings = Column(
            expand = True,
            visible = True,
            controls = [
                Text("設定考題", size = 20, weight = "bold"),
                Row(
                    controls = [
                        self.choose_test_sheetname_text,
                        self.choose_test_sheetname,
                    ],
                ),
                self.test_start_end,
                self.test_submit,
            ]
        )

    # 區塊二：測驗開始
        self.question_information = Column(
            expand = True,
            scroll = "AUTO",
            controls = [],
        )
        self.question_information_up_down_button = Row(
            controls = [],
        )
        self.question_information_main = Column(
            expand = True,
            visible = False,
            controls = [
                self.question_information,
                Divider(height=1),
                self.question_information_up_down_button,
            ],
        )
    # 區塊三：測驗結果
        self.check_result = Column(
            expand = True,
            scroll = "AUTO",
            controls = [],
        )
        self.check_result_main = Column(
            expand = True,
            visible = False,
            controls = [
                Text("測驗結果：", size = 20, weight = "bold"),
                self.check_result,
            ],
        )

    def view(self, page: Page, params: Params, basket: Basket):
        return View(
            controls = [
                Row(
                    expand = True,
                    controls = [
                        Row(
                            width = 100,
                            controls = [rail_item(page)],
                        ),
                        VerticalDivider(width=1),
                        Column(
                            expand = True,
                            alignment = MainAxisAlignment.START,
                            controls = [
                                Row(
                                    controls = [
                                        Text("題庫測驗", size=30, weight="bold"),
                                        self.test_restart,
                                        self.get_results,
                                        self.open_question_navigation,
                                    ],
                                ),
                                Divider(height=1),
                                self.exam_settings,
                                self.question_information_main,
                                self.check_result_main,
                            ],
                        ),
                        VerticalDivider(width=1),
                        self.question_navigation,
                    ],
                ),
            ],
        )