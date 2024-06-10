import flet as ft
import json
import ctypes
from flet import *
from flet_route import Routing, path
from assets.pages.home import Home
from assets.pages.review import Review
from assets.pages.test import Test
from assets.pages.history import History
from assets.pages.settings import Settings

def main(page:ft.Page) -> ft.Page:
    jsonfile = "./assets/data/settings.json"
    with open(jsonfile, encoding="utf-8") as f: # 開啟json檔並讀取設定檔
        read_file = f.read()
        data = json.loads(read_file)
    factor = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100 # 解析度縮放
    screenx = int(ctypes.windll.user32.GetSystemMetrics(0)) # 螢幕寬度
    screeny = int(ctypes.windll.user32.GetSystemMetrics(1)) # 螢幕高度
    backup_window_size = data['window_size'].split('x') # 保留的視窗大小
    page.window_width = int(backup_window_size[0]) / factor
    page.window_height = int(backup_window_size[1]) / factor
    page.window_left = int((screenx - page.window_width)/2)
    page.window_top = int((screeny - page.window_height)/2)
    page.title = "證照考試測驗"

    # 字體
    page.fonts = {"msjhbd": "./assets/fonts/msjhbd.ttc"}

    # Disable Animation Transition 停用動畫過渡
    theme = Theme()
    theme.page_transitions.windows = PageTransitionTheme.NONE
    page.theme = theme
    page.theme_mode = f"{data['theme']}" # 主題
    page.theme = Theme(color_scheme=ColorScheme(primary=data['button_text_color']), font_family = "msjhbd") # 基本元件顏色
    page.update()

    app_routes = [
        path(
            url = "/",
            clear = True,
            view = Home(page).view,
        ),
        path(
            url = "/review",
            clear = True,
            view = Review(page).view,
        ),
        path(
            url = "/test",
            clear = True,
            view = Test(page).view,
        ),
        path(
            url = "/check_history",
            clear = True,
            view = History(page).view,
        ),
        path(
            url = "/settings",
            clear = True,
            view = Settings(page).view,
        ),
    ]
    Routing(
        page = page,
        app_routes = app_routes,
    )
    page.go(page.route)
    page.update()
ft.app(target = main, assets_dir="assets")
