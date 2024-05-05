from flet import *
from flet_route import Params, Basket
from assets.pages.navigation import rail_item

class Settings(UserControl):
    def __init__(self, page):
        super().__init__()
        def change_theme(e):
            if self.theme_switch.value == True:
                page.theme_mode = "LIGHT"
                self.theme_switch.label = "淺色"
            else:
                page.theme_mode = "DARK"
                self.theme_switch.label = "深色"

            page.update()

        self.theme_text = Text("主題：", size=22, weight="bold")
        self.theme_switch = Switch(label="深色", value=False, on_change = change_theme)

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
                                Row(
                                    controls = [
                                        self.theme_text,
                                        self.theme_switch,
                                    ],
                                ),
                            ],
                        ),
                    ],
                ),
            ]
        )