from customtkinter import CTk, CTkTabview, set_appearance_mode, set_default_color_theme

from insert_tab import build_insert_tab
from search_tab import build_search_tab


def build_app():
	set_appearance_mode("System")
	set_default_color_theme("blue")

	app = CTk()
	app.title("Excel Form")
	app.geometry("1200x700")

	tabview = CTkTabview(app)
	tabview.pack(fill="both", expand=True, padx=16, pady=16)

	tab_insert = tabview.add("Insert")
	tab_search = tabview.add("Search")

	build_insert_tab(tab_insert)
	build_search_tab(tab_search)
	tab_insert.refresh_search = getattr(tab_search, "refresh_search", None)

	last_tab = {"name": None}

	def watch_tab_selection():
		current_tab = tabview.get()
		if current_tab != last_tab["name"]:
			last_tab["name"] = current_tab
			if current_tab == "Search" and hasattr(tab_search, "refresh_search"):
				tab_search.refresh_search()
		app.after(200, watch_tab_selection)

	watch_tab_selection()
	return app


if __name__ == "__main__":
	build_app().mainloop()
