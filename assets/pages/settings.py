from flet import *
from flet_route import Params, Basket
from flet_contrib.color_picker import ColorPicker
from assets.pages.navigation import rail_item
import ctypes
import json
import time

class Settings(UserControl):
    def __init__(self, page):
        # super().__init__()
        screenx = int(ctypes.windll.user32.GetSystemMetrics(0)) # 螢幕寬度
        screeny = int(ctypes.windll.user32.GetSystemMetrics(1)) # 螢幕高度
        factor = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100 # 解析度縮放
        print(screenx*factor)
        print(screeny*factor)
        jsonfile = "./assets/data/settings.json" # json檔名

        # 讀取json檔
        def read_json():
            with open(jsonfile, encoding="utf-8") as f: # 開啟json檔並讀取
                read_file = f.read()
                data = json.loads(read_file)
            return data
        
        # 寫入設定元件
        def write_settings_widget():
            data = read_json()
            # print(data)
            if data['theme'] == "LIGHT":
                self.theme_switch.value = True
                self.theme_switch.label = "淺色"
            elif data['theme'] == "DARK":
                self.theme_switch.value = False
                self.theme_switch.label = "深色"
            else: # 待製作套用"系統"顏色
                self.theme_switch.disabled = True
            self.window_size.hint_text = f"{data['window_size']}"
            if data['button_text_color'] != "":
                self.colorpicker.color = data['button_text_color']

        # 保存設定檔至json
        def save_settings(edit_title, edit_content):
            data = read_json()
            data[f"{edit_title}"] = f"{edit_content}"
            with open(jsonfile, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4, ensure_ascii = False)

        # 更改主題背景
        def change_theme(e):
            if self.theme_switch.value == True:
                page.theme_mode = "LIGHT"
                self.theme_switch.label = "淺色"
                save_settings("theme", "LIGHT")
            else:
                page.theme_mode = "DARK"
                self.theme_switch.label = "深色"
                save_settings("theme", "DARK")

            page.update()
        self.theme_switch = Switch(label="深色", value=False, on_change = change_theme)

        # 更改視窗大小
        def change_window_size(e):
            data = read_json()
            self.check_time = 5 # 等待響應時間
            self.accept = False # 是否確認更改畫面大小
            backup_window_size = data['window_size'].split('x') # 保留未修改的螢幕寬度
            size = self.window_size.value.split('x')
            # print(factor)
            page.window_width = int(size[0]) / factor # 寬
            page.window_height = int(size[1]) / factor # 高
            page.window_left = round((screenx - page.window_width)/2)
            page.window_top = round((screeny - page.window_height)/2)
            page.dialog = change_window_size_check
            change_window_size_check.open = True
            page.update()
            while self.check_time and not self.accept: # 倒數並確認響應
                change_window_size_check.content.value = f"是否變更畫面大小? {self.check_time}s"
                # print(self.check_time)
                page.update()
                self.check_time -= 1
                time.sleep(1)
            change_window_size_check.open = False
            page.update()
            if self.check_time == 0 or self.accept == False: # 如果沒響應則返回修改前
                self.window_size.value = ""
                self.window_size.hint_text = f"{data['window_size']}"
                page.window_width = int(backup_window_size[0]) / factor
                page.window_height = int(backup_window_size[1]) / factor
                page.window_left = round((screenx - page.window_width)/2)
                page.window_top = round((screeny - page.window_height)/2)
            page.update()
        def close_dlg(e): # 同意修改
            self.accept = True
            # print(self.window_size.value)
            save_settings("window_size", self.window_size.value)
        def back_size(e): # 不同意修改
            self.accept == False
            self.check_time = 0
        # 修改後二步確認
        change_window_size_check = AlertDialog(
            modal = True,
            title = Text("二步確認"),
            content = Text("是否變更畫面大小?"),
            actions = [
                TextButton("是", on_click = close_dlg),
                TextButton("否", on_click = back_size),
            ],
            actions_alignment = MainAxisAlignment.END,
        )
        self.window_size = Dropdown(
            expand = True,
            label="視窗大小",
            hint_text="None",
            options=[],
            autofocus=True,
            on_change = change_window_size
        )
        for w_size in [
            "800x600",
            "1000x700",
            "1024x768",
            "1280x720",
            "1280x1440",
            "1440x1080",
            "1600x900",
            "1600x1200",
            "1920x1080",
            "1920x1200",
            "1920x1440",
            "2560x1440"
            ]: # 判斷您的螢幕最大大小
            width_height_size = w_size.split('x')
            if int(width_height_size[0]) <= screenx*factor and int(width_height_size[1]) <= screeny*factor:
                self.window_size.options.append(dropdown.Option(f"{w_size}"))
        

        # 設定按鈕文字顏色
        self.colorpicker = ColorPicker()
        def click_colorpicker_submit(e):
            # print(self.colorpicker.color)
            page.theme = Theme(color_scheme=ColorScheme(primary=f"{self.colorpicker.color}"), font_family = "msjhbd") # 調整基本元件顏色
            save_settings("button_text_color", self.colorpicker.color)
            page.update()
        def click_colorpicker_initial(e):
            page.theme = Theme(color_scheme=ColorScheme(primary=""), font_family = "msjhbd") # 重置基本元件顏色
            save_settings("button_text_color", "")
            page.update()
        self.colorpicker_submit = TextButton(
            content = Container(
                content = Column(
                    controls = [
                        Text(value="修改顏色", size=20),
                    ],
                    alignment = MainAxisAlignment.CENTER,
                    spacing = 5,
                ),
                padding = padding.all(10),
            ),
            on_click = click_colorpicker_submit
        )
        self.colorpicker_initial = TextButton(
            content = Container(
                content = Column(
                    controls = [
                        Text(value="重置顏色", size=20),
                    ],
                    alignment = MainAxisAlignment.CENTER,
                    spacing = 5,
                ),
                padding = padding.all(10),
            ),
            on_click = click_colorpicker_initial
        )
        write_settings_widget()

    def view(self, page: Page, params: Params, basket: Basket):
        return View(
            controls = [
                Row(
                    expand = True,
                    controls = [
                        Row(
                            width = 100,
                            controls = [rail_item(page, 0)],
                        ),
                        VerticalDivider(width=1),
                        Column(
                            expand = True,
                            alignment=MainAxisAlignment.START,
                            controls = [
                                Text("設定", size=30, weight="bold"),
                                Divider(height=1),
                                Column(
                                    expand = True,
                                    scroll = "AUTO",
                                    controls = [
                                        Divider(height = 1),
                                        Row(
                                            controls = [
                                                Text("主題：", size=22, weight="bold"),
                                                self.theme_switch,
                                            ],
                                        ),
                                        Divider(height = 1),
                                        Row(
                                            controls = [
                                                Text("視窗大小：", size=22, weight="bold"),
                                                self.window_size,
                                            ],
                                        ),
                                        Divider(height = 1),
                                        Row(
                                            scroll = "AUTO",
                                            controls = [
                                                Text("按鈕文字顏色：", size=22, weight="bold"),
                                                Column(
                                                    controls = [
                                                        self.colorpicker,
                                                        Row(
                                                            controls = [
                                                                self.colorpicker_submit,
                                                                self.colorpicker_initial,
                                                            ],
                                                        ),
                                                    ],
                                                ),
                                            ],
                                        ),
                                        Divider(height = 1),
                                    ],
                                ),
                            ],
                        ),
                    ],
                ),
            ]
        )