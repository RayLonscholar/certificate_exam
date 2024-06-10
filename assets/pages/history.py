from flet import *
from flet_route import Params, Basket
from assets.pages.navigation import rail_item
import openpyxl
import json

class History():
    def __init__(self, page):
        super().__init__()
        def warning_info(text:str):
            info = AlertDialog(
                title=Text(text)
            )
            page.dialog = info
            info.open = True
            page.update()

        def date_list_clicked(e):
            try:
                # print(e.control.content.content.controls[0].value)
                click_data = self.data[f"{e.control.content.content.controls[0].value}"]
                self.history_check_result_main.controls[0].value = f"測驗題庫：{click_data['題庫名稱']}；測驗結果：{round(100 * (click_data['答對題數']/len(click_data['所有題目'])))}分 | 答對{click_data['答對題數']}/{len(click_data['所有題目'])}題 | 錯{len(click_data['所有題目'])-click_data['答對題數']}題"
                self.history_check_result.controls = [] # 清空元件
                # 創建題目清單
                for index, question in enumerate(click_data['所有題目']):
                    order = question[0].split("!n")
                    # 題目
                    if ''.join(click_data['所選的答案'][index]) == question[2].strip( ): # 判斷是否答錯
                        self.history_check_result.controls.append(
                            Row(
                                controls = [
                                    Icon(name=icons.CIRCLE_OUTLINED, color="green"),
                                    Text("{}.({})：{}".format(index+1, question[2], order[0][4:]), size = 20, weight = "bold", expand = True, selectable = True),
                                ],
                            )
                        )
                    else:
                        self.history_check_result.controls.append(
                            Row(
                                controls = [
                                    Icon(name=icons.CLOSE, color="red"),
                                    Text("{}.({})：{}".format(index+1, question[2], order[0][4:]), size = 20, weight = "bold", expand = True, selectable = True),
                                ],
                            )
                        )
                    # 題目圖片
                    for index2, _ in enumerate(eval(question[1])):
                        if _ != "": # 有題目圖片
                            self.history_check_result.controls.append(
                                Image(src=f"{_}")
                            )
                        else: # 沒有題目圖片
                            self.history_check_result.controls.append(
                                Text(f"{order[index2+1]}", size = 20, weight = "bold", selectable = True)
                            )
                    # if question[1] != None:
                    #     self.history_check_result.controls.append(
                    #         Image(src=f"{question[1]}")
                    #     )
                    # for _ in order[1:]:
                    #     self.history_check_result.controls.append(
                    #     Text(f"{_}", size = 20, weight = "bold")
                    #     )
                    # 選項
                    for index3, _ in enumerate(question[3:]):
                        option = ['A', 'B', 'C', 'D']
                        if _ != None:
                            if _[0] in option: # 判斷是否有圖片
                                self.history_check_result.controls.append(
                                    Text(f"{_}", size = 20, weight = "bold")
                                )
                            else:
                                self.history_check_result.controls.append(
                                    Row(
                                        controls = [
                                            Text(f"{option[index3]}：", size = 20, weight = "bold"),
                                            Image(src=f"{_}")
                                        ],
                                    )
                                )
                    # 您所選的答案
                    self.history_check_result.controls.append(
                        Text(f"您所選的答案：{''.join(click_data['所選的答案'][index])}", size = 20, weight = "bold", bgcolor = colors.SECONDARY_CONTAINER)
                    )
                    # 分隔題目
                    self.history_check_result.controls.append(
                        Text("", size = 20, weight = "bold")
                    )
                
                self.date.visible = False # 隱藏第一區塊
                self.history_check_result_main.visible = True # 顯示第二區塊
            except TypeError:
                # print("TypeERROR")
                self.date.visible = True # 顯示第一區塊
                self.history_check_result_main.visible = False # 隱藏第二區塊
                warning_info("無法讀取舊版資料")
            except SyntaxError:
                # print("SyntaxERROR")
                warning_info("檔案資料不完整，缺少圖片等文件檔案")
            page.update()

        # 重新讀取歷史紀錄程序
        def reload_history_click(e):
            self.date_list.controls = [] # 清空元件
            jsonfile = "./assets/data/history.json"
            with open(jsonfile, encoding="utf-8") as f: # 開啟json檔並讀取
                read_file = f.read()
                self.data = json.loads(read_file)
                # print(list(self.data.keys()))
            for index in range(len(self.data)):
                _ = list(self.data.keys())[-(index+1)]
                # print(_)
                self.date_list.controls.append(
                    TextButton(
                        content = Container(
                            content = Column(
                                controls = [
                                    Text(value = "{}".format(_),data = index, size = 18),
                                ],
                                alignment = MainAxisAlignment.CENTER,
                            ),
                        ),
                        on_click = lambda e:date_list_clicked(e)
                    )
                )
            self.date.visible = True # 顯示第一區塊
            self.history_check_result_main.visible = False # 隱藏第二區塊
            page.update()

        # 重新讀取歷史紀錄按鈕
        self.reload_history = TextButton(
            content = Container(
                content=Column(
                    controls = [
                        Text(value="重新讀取歷史資料", size=20),
                    ],
                    alignment=MainAxisAlignment.CENTER,
                    spacing=5,
                ),
            ),
            on_click=reload_history_click)

        # 區塊1：考試日期按鈕清單
        # 題目按鈕清單
        self.date_list = Column(
            controls = [],
        )
        # 主區塊
        self.date = Column(
            expand = True,
            visible = True,
            controls = [
                Text("選擇測驗日期", size = 20, weight = "bold"),
                Column(
                    expand = True,
                    scroll = "AUTO",
                    controls = [self.date_list],
                ),
            ],
        )

        # 第二區塊
        self.history_check_result = Column(
            expand = True,
            scroll = "AUTO",
            controls = [],
        )
        # 主區塊
        self.history_check_result_main = Column(
            expand = True,
            visible = False,
            controls = [
                Text("測驗題庫：", size = 20,weight = "bold"),
                Divider(height=1),
                self.history_check_result,
            ],
        )

        reload_history_click(self)
    
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
                            alignment=MainAxisAlignment.START,
                            controls = [
                                Row(
                                    controls = [
                                        Text("查看歷史測驗結果", size=30, weight="bold"),
                                        self.reload_history,
                                    ]
                                ),
                                Divider(height=1),
                                self.date,
                                self.history_check_result_main,
                            ],
                        ),
                    ],
                ),
            ],
        )