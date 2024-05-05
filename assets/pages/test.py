from flet import *
from flet_route import Params, Basket
from assets.pages.navigation import rail_item
import openpyxl

class Test(UserControl):
    def __init__(self, page):
        super().__init__()

        # 選擇題庫(使用review)
        # 輸入考題數

        self.exam_settings = Column(
            expand = True,
            controls = [
                Text("設定考題", size = 20, weight = "bold"),
            ],
        )
        
        def test_clicked(e):
            page.go("/test")

        self.test = TextButton(
            content = Container(
                content = Column(
                    controls = [
                        Text(value="題庫測驗", size=20),
                    ],
                    alignment = MainAxisAlignment.CENTER,
                    spacing = 5,
                ),
                padding = padding.all(10),
            ),
            on_click = test_clicked)

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
                                Text("題庫測驗", size=30, weight="bold"),
                                Divider(height=1),
                                self.exam_settings,
                                self.test,
                            ],
                        ),
                    ],
                ),
            ],
        )