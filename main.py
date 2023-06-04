from logging import log
import folium
import pandas as pd
import webbrowser


def make_html_file():
    franchises = pd.read_excel("aryta.xlsx", sheet_name="1")
    center = [35.751355, 51.386032]
    map_aryta = folium.Map(location=center, zoom_start=12)
    i = 0
    for index, franchise in franchises.iterrows():
        i += 1
        try:
            location = [franchise["lat"], franchise["long"]]
            folium.Marker(
                location, popup=f'{franchise["name"]} {franchise["address"]} '
            ).add_to(map_aryta)
            folium.CircleMarker(
                location=location, fill_color="#3186cc", radius=20, color="#3186cc"
            ).add_to(map_aryta)
        except:
            print(f"Error in reading line number {i}")
    map_aryta.save("aryta_locations.html")


def main():
    make_html_file()
    webbrowser.open_new_tab("aryta_locations.html")


if __name__ == "__main__":
    main()
