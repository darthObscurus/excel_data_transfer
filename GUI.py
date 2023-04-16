import os
import winreg

from tkinter import *
from tkinter import ttk, filedialog

from main import get_data_from, write_data_to


def get_user_accent_color():
    light_accent_color = ''
    lighter_accent_color = ''

    def get_color_from_reg():
        reg_key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            'SOFTWARE\Microsoft\Windows\DWM',
            access=winreg.KEY_READ | winreg.KEY_WOW64_32KEY
        )
        value, type = winreg.QueryValueEx(reg_key, "ColorizationAfterglow")
        color = hex(value)[2:]
        accent_color = color[2:]
        return color, accent_color

    def get_same_opaque_color(argb):
        return (int(255 - argb['a'] * (255 - argb['r'])),
                int(255 - argb['a'] * (255 - argb['g'])),
                int(255 - argb['a'] * (255 - argb['b'])))

    argb_color, accent_color = get_color_from_reg()
    argb = {
        'a': argb_color[:2],
        'r': argb_color[2:4],
        'g': argb_color[4:6],
        'b': argb_color[6:],
    }

    for name, hex_code in argb.items():
        argb[name] = (int(hex_code, 16))
    argb['a'] = round(argb['a'] / 255, 3)

    rgb = get_same_opaque_color(argb)
    for color in rgb:
        light_accent_color += hex(color)[-2:]

    argb['a'] = 0.627
    lighter_accent_rgb = get_same_opaque_color(argb)
    for color in lighter_accent_rgb:
        lighter_accent_color += hex(color)[-2:]

    return accent_color, light_accent_color, lighter_accent_color


root = Tk()

accent_color, light_accent_color, lighter_accent_color = get_user_accent_color()
bg_color = "#e6e6e6"
text_color = "black"
style = ttk.Style()
print(style.theme_names())
style.theme_use("clam")
style.configure("frame.TFrame",
                background=bg_color)
style.configure("label.TLabel",
                background=bg_color,
                foreground=text_color,
                font=("Segoe UI", 12),
                )
style.configure("label_emoji.TLabel",
                background=bg_color,
                # foreground="default",
                font=("Segoe UI Emoji", 14),
                )
style.configure("label_success.TLabel",
                background=bg_color,
                foreground="green",
                font=("Segoe UI Emoji", 14),
                )
style.configure("btn.TButton",
                background="#fafafa",
                borderwidth=2,
                padding=[-10, 15, -10, 15],
                relief="flat",
                font=("Segoe UI", 14, "bold"),
                foreground=text_color
                )
style.map("btn.TButton",
          background=[("pressed", "#d1d1d1"), ("active", "#ddd"),
                          ("disabled", "#f0f0f0")],
          focuscolor="#fafafa",
          focus="clear",
          relief=[("pressed", "flat"), ("focus", "solid"), ],
          bordercolor=[("focus", f"#{accent_color}")],
          )
style.configure("btn_accent.TButton",
                background=f"#{lighter_accent_color}",
                borderwidth=2,
                padding=[-10, 15, -10, 15],
                relief="flat",
                font=("Segoe UI", 14, "bold"),
                foreground=text_color
                )
style.map("btn_accent.TButton",
          background=[("pressed", f"#{accent_color}"), ("active", f"#{light_accent_color}"),
                          ("disabled", "#f0f0f0")],
          focuscolor="#fafafa",
          focus="clear",
          relief=[("pressed", "flat"), ("focus", "solid"), ],
          bordercolor=[("focus", f"#{accent_color}")],
          foreground=[("disabled", "#999999")]
          )
style.configure("btn_red.TButton",
                background="#fafafa",
                borderwidth=2,
                padding=[-10, 15, -10, 15],
                relief="flat",
                font=("Segoe UI", 14, "bold"),
                foreground=text_color
                )
style.map("btn_red.TButton",
          background=[("pressed", "#d1d1d1"), ("active", "#ddd"),
                          ("disabled", "#f0f0f0")],
          focuscolor="#fafafa",
          focus="clear",
          relief=[("pressed", "flat"), ("focus", "solid"), ],
          bordercolor=[("focus", f"#{accent_color}")],
          foreground=[("active", "red"), ("disabled", "#999999")]
          )

style.configure("entry.TEntry",
                fieldbackground="#f0f0f0",
                bordercolor="#f0f0f0",
                selectbackground = f"#{light_accent_color}",
                foreground=text_color
                )
style.map("entry.TEntry",
          lightcolor=[("focus", f"#{accent_color}")],
          bordercolor=[("focus", f"#{accent_color}")]
          )


reading_books = {
    "ДНС-8": "",
    "УПВСН": "",
    "ДНС-1": "",
    "ДНС-210": "",
    "ГЗНУ-4304": ""
}
filename = ""


def open_file(object_name):
    filepath = filedialog.askopenfilename(defaultextension="xls")
    if filepath != "":
        reading_books[object_name] = os.path.abspath(filepath)
        if object_name == "ДНС-8":
            path_DNS_8.set(reading_books[object_name])
            clear_btn_DNS_8.state(["!disabled"])
            path, filename = os.path.split(filepath)
        elif object_name == "УПВСН":
            path_UPVSN.set(reading_books[object_name])
            clear_btn_UPVSN.state(["!disabled"])
            path, filename = os.path.split(filepath)
        elif object_name == "ДНС-1":
            path_DNS_1.set(reading_books[object_name])
            clear_btn_DNS_1.state(["!disabled"])
            path, filename = os.path.split(filepath)
        elif object_name == "ДНС-210":
            path_DNS_210.set(reading_books[object_name])
            clear_btn_DNS_210.state(["!disabled"])
            path, filename = os.path.split(filepath)
        elif object_name == "ГЗНУ-4304":
            path_GZNU.set(reading_books[object_name])
            clear_btn_GZNU.state(["!disabled"])
            path, filename = os.path.split(filepath)
        print(reading_books)
        start.state(["!disabled"])


def clear_path(object_name):
    global reading_book
    if object_name == "ДНС-8":
        path_DNS_8.set("Файл не выбран")
        clear_btn_DNS_8.state(["disabled"])
    elif object_name == "УПВСН":
        path_UPVSN.set("Файл не выбран")
        clear_btn_UPVSN.state(["disabled"])
    elif object_name == "ДНС-1":
        path_DNS_1.set("Файл не выбран")
        clear_btn_DNS_1.state(["disabled"])
    elif object_name == "ДНС-210":
        path_DNS_210.set("Файл не выбран")
        clear_btn_DNS_210.state(["disabled"])
    elif object_name == "ГЗНУ-4304":
        path_GZNU.set("Файл не выбран")
        clear_btn_GZNU.state(["disabled"])
    reading_books[object_name] = ""
    print(reading_books)
    is_start_diabled = True
    for book in reading_books.values():
        if book != "":
            is_start_diabled = False
    if is_start_diabled:
        start.state(["disabled"])


def transfer_data():
    for object_index in range(len(reading_books)):
        main_frame.winfo_children()[object_index].winfo_children()[-1]["text"] = \
            "💤"
        main_frame.winfo_children()[object_index].winfo_children()[-1].update()
    index = 0
    for obj_name, book in reading_books.items():
        main_frame.winfo_children()[index].winfo_children()[-1]["text"] = \
            "📤"
        main_frame.winfo_children()[index].winfo_children()[-1].update()
        if book != '':
            sorted_data = get_data_from(book)
            write_data_to(sorted_data, obj_name)
        main_frame.winfo_children()[index].winfo_children()[-1]["text"] = \
            "✔"
        main_frame.winfo_children()[index].winfo_children()[-1]["style"] = \
            "label_success.TLabel"
        main_frame.winfo_children()[index].winfo_children()[-1].update()
        index += 1


path_DNS_8 = StringVar(value="Файл не выбран")
path_UPVSN = StringVar(value="Файл не выбран")
path_DNS_1 = StringVar(value="Файл не выбран")
path_DNS_210 = StringVar(value="Файл не выбран")
path_GZNU = StringVar(value="Файл не выбран")
main_frame = ttk.Frame(root, padding=5, style="frame.TFrame")
main_frame.grid()
DNS_8_frame = ttk.Frame(main_frame, name="dns_8", style="frame.TFrame")
DNS_8_frame.grid(column=0, row=0, columnspan=4)
UPVSN_frame = ttk.Frame(main_frame, name="upvsn", style="frame.TFrame")
UPVSN_frame.grid(column=0, row=1, columnspan=4)
DNS_1_frame = ttk.Frame(main_frame, name="dns_1", style="frame.TFrame")
DNS_1_frame.grid(column=0, row=2, columnspan=4)
DNS_210_frame = ttk.Frame(main_frame, name="dns_210", style="frame.TFrame")
DNS_210_frame.grid(column=0, row=3, columnspan=4)
GZNU_frame = ttk.Frame(main_frame, name="gznu", style="frame.TFrame")
GZNU_frame.grid(column=0, row=4, columnspan=4)

# Блок ДНС-8
object_name = ttk.Label(master=DNS_8_frame, text="ДНС-8", style="label.TLabel")
object_name.grid(column=0, row=0, sticky="w")
DNS_8_path = ttk.Entry(DNS_8_frame, textvariable=path_DNS_8, width=30,
                       style="entry.TEntry", font=("Segoe UI", 14, "bold"))
DNS_8_path.grid(column=0,
                row=1,
                columnspan=3,
                sticky="w",
                padx=1,
                pady=1,
                ipady=14)
browse_btn = ttk.Button(DNS_8_frame,
                        text="👀",
                        command=lambda a="ДНС-8": open_file(a),
                        style="btn.TButton")
browse_btn.grid(column=3, row=1, padx=1, pady=1)
clear_btn_DNS_8 = ttk.Button(DNS_8_frame,
                             text="❌",
                             command=lambda a="ДНС-8": clear_path(a),
                             state="disabled",
                             style="btn_red.TButton"
                             )
clear_btn_DNS_8.grid(column=4, row=1, padx=1, pady=1)
msg_label = ttk.Label(DNS_8_frame, name="dns_8_msg",
                      style="label_emoji.TLabel")
msg_label.grid(column=3, row=4)


# Блок УПВСН
object_name = ttk.Label(UPVSN_frame, text="УПВСН", style="label.TLabel")
object_name.grid(column=0, row=0, sticky="w")
UPVSN_path = ttk.Entry(UPVSN_frame, textvariable=path_UPVSN, width=30, style="entry.TEntry",
                       font=("Segoe UI", 14, "bold"))
UPVSN_path.grid(column=0,
                row=1,
                columnspan=3,
                sticky="w",
                padx=1,
                pady=1,
                ipady=14)
browse_btn = ttk.Button(UPVSN_frame,
                        text="👀",
                        command=lambda a="УПВСН": open_file(a),
                        style="btn.TButton")
browse_btn.grid(column=3, row=1, padx=1, pady=1)
clear_btn_UPVSN = ttk.Button(UPVSN_frame,
                             text="❌",
                             command=lambda a="УПВСН": clear_path(a),
                             state="disabled",
                             style="btn_red.TButton")
clear_btn_UPVSN.grid(column=4, row=1, padx=1, pady=1)
msg_label = ttk.Label(UPVSN_frame, style="label_emoji.TLabel")
msg_label.grid(column=3, row=4)


# Блок ДНС-1
object_name = ttk.Label(DNS_1_frame, text="ДНС-1", style="label.TLabel")
object_name.grid(column=0, row=0, sticky="w")
DNS_1_path = ttk.Entry(DNS_1_frame, textvariable=path_DNS_1, width=30, style="entry.TEntry",
                       font=("Segoe UI", 14, "bold"))
DNS_1_path.grid(column=0,
                row=1,
                columnspan=3,
                sticky="w",
                padx=1,
                pady=1,
                ipady=14)
browse_btn = ttk.Button(DNS_1_frame,
                        text="👀",
                        command=lambda a="ДНС-1": open_file(a),
                        style="btn.TButton")
browse_btn.grid(column=3, row=1, padx=1, pady=1)
clear_btn_DNS_1 = ttk.Button(DNS_1_frame,
                             text="❌",
                             command=lambda a="ДНС-1": clear_path(a),
                             state="disabled",
                             style="btn_red.TButton")
clear_btn_DNS_1.grid(column=4, row=1, padx=1, pady=1)
msg_label = ttk.Label(DNS_1_frame, style="label_emoji.TLabel")
msg_label.grid(column=3, row=4)


# Блок ДНС-210
object_name = ttk.Label(DNS_210_frame, text="ДНС-210", style="label.TLabel")
object_name.grid(column=0, row=0, sticky="w")
DNS_210_path = ttk.Entry(DNS_210_frame, textvariable=path_DNS_210, width=30, style="entry.TEntry",
                         font=("Segoe UI", 14, "bold"))
DNS_210_path.grid(column=0,
                  row=1,
                  columnspan=3,
                  sticky="w",
                  padx=1,
                  pady=1,
                  ipady=14)
browse_btn = ttk.Button(DNS_210_frame,
                        text="👀",
                        command=lambda a="ДНС-210": open_file(a),
                        style="btn.TButton")
browse_btn.grid(column=3, row=1, padx=1, pady=1)
clear_btn_DNS_210 = ttk.Button(DNS_210_frame,
                               text="❌",
                               command=lambda a="ДНС-210": clear_path(a),
                               state="disabled",
                               style="btn_red.TButton")
clear_btn_DNS_210.grid(column=4, row=1, padx=1, pady=1)
msg_label = ttk.Label(DNS_210_frame, style="label_emoji.TLabel")
msg_label.grid(column=3, row=4)


# Блок ГЗНУ-4304
object_name = ttk.Label(GZNU_frame, text="ГЗНУ-4304", style="label.TLabel")
object_name.grid(column=0, row=0, sticky="w")
GZNU_path = ttk.Entry(GZNU_frame, textvariable=path_GZNU, width=30, style="entry.TEntry",
                      font=("Segoe UI", 14, "bold"))
GZNU_path.grid(column=0,
               row=1,
               columnspan=3,
               sticky="w",
               padx=1,
               pady=1,
               ipady=14)
browse_btn = ttk.Button(GZNU_frame,
                        text="👀",
                        command=lambda a="ГЗНУ-4304": open_file(a),
                        style="btn.TButton")
browse_btn.grid(column=3, row=1, padx=1, pady=1)
clear_btn_GZNU = ttk.Button(GZNU_frame,
                            text="❌",
                            command=lambda a="ГЗНУ-4304": clear_path(a),
                            state="disabled",
                            style="btn_red.TButton")
clear_btn_GZNU.grid(column=4, row=1, padx=1, pady=1)
msg_label = ttk.Label(GZNU_frame, style="label_emoji.TLabel")
msg_label.grid(column=3, row=4, pady=5)


start = ttk.Button(main_frame,
                   text="Начать",
                   command=transfer_data,
                   state="disabled", style="btn_accent.TButton")
start.grid(column=0, row=10, columnspan=2, sticky="nsew", padx=1, pady=1)
ttk.Button(main_frame,
           text="Выйти",
           command=root.destroy, style="btn.TButton").grid(column=2, row=10, columnspan=2,
                                       sticky="nsew", padx=1, pady=1)

root.mainloop()