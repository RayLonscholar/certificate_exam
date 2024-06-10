from flet import *

# NavigationBar
def bar_item(page: Page, now_selected: int = None):
    def change_page(e, page):
        print(page.route)
        print(e.control.selected_index)
        if e.control.selected_index == 0:
            page.go("/")

        if e.control.selected_index == 1:
            page.go("/about")
        page.update()

    nav_bar_item = NavigationBar(
        on_change = lambda e:change_page(e, page),
        destinations = [
            NavigationDestination(icon = icons.HOME, label = "home"),
            NavigationDestination(icon = icons.EXPLORE, label = "about"),
        ],
        selected_index = now_selected,
    )
    return nav_bar_item

# NavigationRail
def rail_item(page: Page, now_selected: int = None):

    def change_page(e, page):
        # print(page.route)
        # print(e.control.selected_index)
        if e.control.selected_index == 0:
            page.go("/settings")
        page.update()

    def back_home(e, page):
        page.go("/")
        page.update()

    rail = NavigationRail(
        selected_index = now_selected, # 選中項
        expand = True,
        label_type = NavigationRailLabelType.ALL,
        min_width = 100,
        min_extended_width = 300,
        leading = FloatingActionButton(icon = icons.HOME, text = "主頁面", on_click = lambda e: back_home(e, page)),
        group_alignment = -0.9,
        destinations = [
            NavigationRailDestination(
                icon=icons.SETTINGS_OUTLINED,
                selected_icon_content=Icon(icons.SETTINGS),
                label_content=Text("設定"),
            ),
        ],
        # on_change=lambda e: print("Selected destination:", e.control.selected_index),
        on_change=lambda e: change_page(e, page)
    )


    return rail