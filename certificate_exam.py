import flet as ft
from flet import *
from urllib.parse import urlparse


def main(page: Page):
	page.title = "routing app"
	youparams = "watermelon"

	def route_change(route):
		# 清除所有視窗
		page.views.clear()
		page.views.append(
			View(
				"/", # 主頁面 URL
				[
				AppBar(title=Text("Home Page", size=30,
					color="white"
					),
					bgcolor="red500",
					),
					# PAGE ROUTE IS PATH YOU URL HERE
					Text(page.route),
					ElevatedButton(
						"Go to Second Page",
						on_click=lambda _: page.go(f"/secondpage/{youparams}")
						)
					]
					)
			)
		# GET PARAM FROM HOME PAGE
		param = page.route
		# THIS IS GET VALUE AFTER /secondpage/THIS RES HERE
		res = urlparse(param).path.split("/")[-1]
		print(f"test res is : {res}")
		if page.route == f"/secondpage/{res}":
			page.views.append(
				View(
				# IF URL ACCESS HERE THEN PUSH TO VIEW HERE
				f"/secondpage/{res}",
				[
					# LIKE BEFORE
					# RENDER YOU PAGE HERE
					AppBar(title=Text("SECOND PAGE", 
						color="white",
						size=30),
						bgcolor="blue500",
						),
					# PAGE ROUTE IS PATH YOU URL HERE
					Text(page.route),
					Text(f"you params is {res}"),
					ElevatedButton(
						"BACK TO HOME PAGE",
						on_click=lambda _: page.go("/")
						)

					]
					)
				)
	page.update()

	def view_pop(view):
		page.views.pop()
		myview = page.views[-1]
		page.go(myview.route)

	page.on_route_change = route_change
	page.on_view_pop = view_pop
	page.go(page.route)
	p = TemplateRoute(page.route)
	if p.match("/second/:id"):
		print("you here ", p.id)
	else:
		print("whatever")


# import flet as ft

# def main(page:ft.Page) -> None:
#     page.title = "證照考試測驗"
#     page.window_width = 1000
#     page.window_height = 700
#     page.bgcolor = "#1f2128"
    
#     def route_change(route):
#         # 清除所有視窗
#         page.views.clear()
#         page.views.append(
#             View(
#         )

#     page.update()

if __name__ == "__main__":
    ft.app(target = main)