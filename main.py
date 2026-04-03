from customtkinter import CTk, CTkFont, CTkTabview, set_appearance_mode, set_default_color_theme

from insert_tab import build_insert_tab
from search_tab import build_search_tab


def build_app():
	set_appearance_mode("System")
	set_default_color_theme("dark-blue")

	app = CTk()
	app.title("Excel Form")
	app.geometry("740x840")

	tabview = CTkTabview(app)
	tabview.pack(fill="both", expand=True, padx=0, pady=10)

	tab_insert = tabview.add("Insert")
	tab_search = tabview.add("Search")

	# Increase tab button size for better click targets and readability.
	segmented_button = getattr(tabview, "_segmented_button", None)
	if segmented_button is not None:
		segmented_button.configure(
			height=40,
			font=CTkFont(size=15, weight="bold"),
			corner_radius=18,
		)
		buttons = getattr(segmented_button, "_buttons_dict", {})
		for button in buttons.values():
			button.configure(width=130, corner_radius=18)

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
