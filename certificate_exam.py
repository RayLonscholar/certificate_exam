import flet as ft
from flet import *
from flet_route import Routing, path
from assets.pages.home import Home
from assets.pages.review import Review
from assets.pages.test import Test
from assets.pages.settings import Settings

def main(page:ft.Page) -> ft.Page:
    page.window_width = 1000
    page.window_height = 700
    page.title = "證照考試測驗"
    page.theme_mode = "DARK"

    # Disable Animation Transition 停用動畫過渡
    theme = Theme()
    theme.page_transitions.windows = PageTransitionTheme.NONE
    page.theme = theme
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
ft.app(target = main)
