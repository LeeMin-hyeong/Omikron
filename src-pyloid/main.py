from pyloid.tray import (
    TrayEvent,
)
from pyloid.utils import (
    get_production_path,
    is_production,
)
from pyloid.serve import pyloid_serve
from pyloid import Pyloid
from server import server
import os

# initiate program directory structure
if not os.path.exists("./data"):
    os.makedirs("./data")
if not os.path.exists("./data/backup"):
    os.makedirs("./data/backup")
os.environ["WDM_PROGRESS_BAR"] = "0"

WIDTH, HEIGHT = 1400, 800

app = Pyloid(app_name="Omikron", single_instance=True, server=server)

app.set_icon(get_production_path("src-pyloid/icons/omikron.png"))
app.set_tray_icon(get_production_path("src-pyloid/icons/omikron.png"))

############################## Tray ################################
def on_double_click():
    app.show_and_focus_main_window()


app.set_tray_actions(
    {
        TrayEvent.DoubleClick: on_double_click,
    }
)
app.set_tray_menu_items(
    [
        {"label": "Show Window", "callback": app.show_and_focus_main_window},
        {"label": "Exit", "callback": app.quit},
    ]
)
####################################################################

if is_production():
    url = pyloid_serve(directory=get_production_path("dist-front"))
    window = app.create_window(
        title="Omikron",
        width=WIDTH,
        height=HEIGHT,
        transparent=True,
    )
    window.load_url(url)
else:
    window = app.create_window(
        title="Omikron-dev",
        dev_tools=True,
        width=WIDTH,
        height=HEIGHT,
        transparent=True,
    )
    window.load_url("http://localhost:5173")

window.set_resizable(False)
window.show_and_focus()

app.run()
