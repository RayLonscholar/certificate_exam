from flet import *
from flet_route import Params, Basket
from assets.pages.navigation import rail_item

class Home(UserControl):
    def __init__(self, page):
        # super().__init__() # 0.20版後則不需要手動繼承
        def review_clicked(e):
            page.go("/review")

        self.review = TextButton(
            content = Container(
                content = Column(
                    controls = [
                        Text(value = "複習題庫題目", size=20),
                    ],
                    alignment = MainAxisAlignment.CENTER,
                    spacing = 5,
                ),
                padding = padding.all(10),
            ),
            on_click = review_clicked)
        
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
        
        def check_history_click(e):
            page.go("/check_history")

        self.check_history = TextButton(
            content = Container(
                content = Column(
                    controls = [
                        Text(value="查看歷史測驗結果", size=20),
                    ],
                    alignment = MainAxisAlignment.CENTER,
                    spacing = 5,
                ),
                padding = padding.all(10),
            ),
            on_click = check_history_click
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
                                Text("主頁面", size=30, weight="bold"),
                                Divider(height=1),
                                self.review,
                                self.test,
                                self.check_history,
                            ],
                        ),
                    ],
                ),
            ],
        )