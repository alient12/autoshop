import win32com.client
import os
import tkinter as tk
from tkinter import ttk

BASE_PATH = os.getcwd() + "/"  # BASE_PATH is an absolute path.
BACKGROUNDS_PATH = BASE_PATH + "backgrounds/"
FOREGROUNDS_PATH = BASE_PATH + "foregrounds/"
OUTPUTS_PATH = BASE_PATH + "outputs/"
MASKS_PATH = BASE_PATH + "masks/"
OUTPUT_IMAGE_WIDTH = 640
OUTPUT_IMAGE_HEIGHT = 480
DEFAULT_BLEND_MODE = "psSubtract"
PsBlendMode = [
    "none",
    "psPassThrough",
    "psNormalBlend",
    "psDissolve",
    "psDarken",
    "psMultiply",
    "psColorBurn",
    "psLinearBurn",
    "psLighten",
    "psScreen",
    "psColorDodge",
    "psLinearDodge",
    "psOverlay",
    "psSoftLight",
    "psHardLight",
    "psVividLight",
    "psLinearLight",
    "psPinLight",
    "psDifference",
    "psExclusion",
    "psHue",
    "psSaturationBlend",
    "psColorBlend",
    "psLuminosity",
    "psHardMix",
    "psLighterColor",
    "psDarkerColor",
    "psSubtract",
    "psDivide",
]

# preper list of foreground images
if not os.path.exists(FOREGROUNDS_PATH):
    raise Exception(
        FOREGROUNDS_PATH
        + " does not exist, please check 'FOREGROUNDS_PATH' and 'BASE_PATH' at the first lines of the autoshop code!"
    )
fore_files = os.listdir(FOREGROUNDS_PATH)
if not len(fore_files):
    raise Exception(
        "There is no image file in "
        + FOREGROUNDS_PATH
        + " , please fill it before using autoshop :)"
    )
fg_path_list = [FOREGROUNDS_PATH + file for file in fore_files]

# preper list of background images
if not os.path.exists(BACKGROUNDS_PATH):
    raise Exception(
        BACKGROUNDS_PATH
        + " does not exist, please check 'BACKGROUNDS_PATH' and 'BASE_PATH' at the first lines of the autoshop code!"
    )
back_files = os.listdir(BACKGROUNDS_PATH)
if not len(back_files):
    raise Exception(
        "There is no image file in "
        + BACKGROUNDS_PATH
        + " , please fill it before using autoshop :)"
    )
bg_path_list = [BACKGROUNDS_PATH + file for file in back_files]

if not os.path.exists(OUTPUTS_PATH):
    awnser = input(
        "\033[38;2;255;255;0m"
        + OUTPUTS_PATH
        + " does not exist, would you like to create it?[y/n]\033[0m"
    )
    if awnser.lower() == "yes" or awnser.lower() == "y" or awnser.lower() == "yeah":
        os.makedirs(OUTPUTS_PATH)
    else:
        raise Exception(
            "\033[38;2;255;0;0mcome back when you made this folder or changed my defaults at first lines of code!\033[0m"
        )

blend_mode = DEFAULT_BLEND_MODE


def open_bg_ps(psApp, bg_index=0):
    psApp.Open(bg_path_list[bg_index])
    doc = psApp.Application.ActiveDocument

    # resize output image
    doc.ResizeImage(
        Width=OUTPUT_IMAGE_WIDTH,
        Height=OUTPUT_IMAGE_HEIGHT,
        Resolution=300,
        ResampleMethod=8,
    )

    return doc


def add_image_layer(psApp, doc, fg_index=0):
    psApp.Load(fg_path_list[fg_index])
    psApp.ActiveDocument.Selection.SelectAll()
    psApp.ActiveDocument.Selection.Copy()
    psApp.ActiveDocument.Close()
    psApp.ActiveDocument.Paste()

    layer = doc.ActiveLayer
    return layer


def replace_image_layer(psApp, doc, layer, fg_index):
    layer.Delete()
    return add_image_layer(psApp, doc, fg_index)


def set_layer_blend_mode(layer, blend_mode):
    layer.BlendMode = PsBlendMode.index(blend_mode)


def set_layer_opacity(layer, value):
    layer.Opacity = value


def add_sold_layer(psApp, doc):
    psApp.Load(BASE_PATH + "solid.png")
    psApp.ActiveDocument.Selection.SelectAll()
    psApp.ActiveDocument.Selection.Copy()
    psApp.ActiveDocument.Close()
    psApp.ActiveDocument.Paste()
    solid_layer = doc.ActiveLayer
    return solid_layer


def save_as(doc, save_path):
    jpgSaveOptions = win32com.client.Dispatch("Photoshop.JPEGSaveOptions")
    doc.SaveAs(
        save_path,
        jpgSaveOptions,
        True,
        2,
    )


class toolbox:
    def __init__(self, bg_index=0, fg_index=0, blend_mode=DEFAULT_BLEND_MODE):
        # connect to photoshop or open it if not opened
        self.psApp = win32com.client.Dispatch("Photoshop.Application")

        self.bg_index = bg_index
        self.fg_index = fg_index

        # create doc with a background image
        self.doc = open_bg_ps(self.psApp, self.bg_index)
        self.bg_layer = self.doc.ArtLayers["Background"]

        self.solid_layer = add_sold_layer(self.psApp, self.doc)
        set_layer_opacity(self.solid_layer, 0)

        # add a foreground image to doc
        self.layer = add_image_layer(self.psApp, self.doc, self.fg_index)

        self.blend_mode = blend_mode

        # set blend mode for layer
        set_layer_blend_mode(self.layer, self.blend_mode)

    def run(self):
        def set_opacity(value):
            set_layer_opacity(self.layer, int(value))

        def set_blend_mode(event):
            self.blend_mode = blend_combo.get()
            self.layer.BlendMode = PsBlendMode.index(self.blend_mode)

        def select_fg(event):
            self.fg_index = fore_files.index(fg_combo.get())
            self.layer = replace_image_layer(
                self.psApp, self.doc, self.layer, self.fg_index
            )
            set_layer_blend_mode(self.layer, self.blend_mode)

        def save():
            save_path = (
                OUTPUTS_PATH
                + back_files[self.bg_index][:-4]
                + "-"
                + fore_files[self.fg_index][:-4]
                + ".jpg"
            )
            save_as(self.doc, save_path)

            set_layer_opacity(self.bg_layer, 0)
            set_layer_opacity(self.solid_layer, 100)

            mask_save_path = (
                MASKS_PATH
                + back_files[self.bg_index][:-4]
                + "-"
                + fore_files[self.fg_index][:-4]
                + ".jpg"
            )

            # self.doc.Export(ExportIn=mask_save_path, ExportAs=2)
            save_as(self.doc, mask_save_path)
            set_layer_opacity(self.bg_layer, 100)
            set_layer_opacity(self.solid_layer, 0)

        def select_bg(event):
            SILENT_CLOSE = 2
            self.doc.Close(SILENT_CLOSE)
            self.bg_index = back_files.index(bg_combo.get())
            self.doc = open_bg_ps(self.psApp, self.bg_index)
            self.bg_layer = self.doc.ArtLayers["Background"]
            self.solid_layer = add_sold_layer(self.psApp, self.doc)
            set_layer_opacity(self.solid_layer, 0)
            self.fg_index = fore_files.index(fg_combo.get())
            self.layer = add_image_layer(self.psApp, self.doc, self.fg_index)
            set_layer_blend_mode(self.layer, self.blend_mode)

        def next_fg():
            if self.fg_index < len(fore_files) - 1:
                self.fg_index += 1
                self.layer = replace_image_layer(
                    self.psApp, self.doc, self.layer, self.fg_index
                )
                set_layer_blend_mode(self.layer, self.blend_mode)
                fg_combo.current(self.fg_index)

        def back_fg():
            if self.fg_index:
                self.fg_index -= 1
                self.layer = replace_image_layer(
                    self.psApp, self.doc, self.layer, self.fg_index
                )
                set_layer_blend_mode(self.layer, self.blend_mode)
                fg_combo.current(self.fg_index)

        def next_bg():
            if self.bg_index < len(back_files) - 1:
                SILENT_CLOSE = 2
                self.doc.Close(SILENT_CLOSE)
                self.bg_index += 1
                self.doc = open_bg_ps(self.psApp, self.bg_index)
                self.bg_layer = self.doc.ArtLayers["Background"]
                self.solid_layer = add_sold_layer(self.psApp, self.doc)
                set_layer_opacity(self.solid_layer, 0)
                self.layer = add_image_layer(self.psApp, self.doc, self.fg_index)
                set_layer_blend_mode(self.layer, self.blend_mode)
                bg_combo.current(self.bg_index)

        def back_bg():
            if self.bg_index:
                SILENT_CLOSE = 2
                self.doc.Close(SILENT_CLOSE)
                self.bg_index -= 1
                self.doc = open_bg_ps(self.psApp, self.bg_index)
                self.bg_layer = self.doc.ArtLayers["Background"]
                self.solid_layer = add_sold_layer(self.psApp, self.doc)
                set_layer_opacity(self.solid_layer, 0)
                self.layer = add_image_layer(self.psApp, self.doc, self.fg_index)
                set_layer_blend_mode(self.layer, self.blend_mode)
                bg_combo.current(self.bg_index)

        master = tk.Tk()
        master.geometry("600x250")
        master.wm_attributes("-topmost", 1)

        tk.Scale(
            master,
            from_=0,
            to=100,
            tickinterval=10,
            length=600,
            orient=tk.HORIZONTAL,
            command=set_opacity,
        ).pack()

        bg_label = tk.Label(master, text="Background:").place(x=5, y=65)

        bg_combo = ttk.Combobox(master, width=69, values=back_files)
        bg_combo.place(x=85, y=65)
        bg_combo.current(self.bg_index)
        bg_combo.bind("<<ComboboxSelected>>", select_bg)

        fg_label = tk.Label(master, text="Foreground:").place(x=5, y=25 + 65)

        fg_combo = ttk.Combobox(master, width=69, values=fore_files)
        fg_combo.place(x=85, y=25 + 65)
        fg_combo.current(self.fg_index)
        fg_combo.bind("<<ComboboxSelected>>", select_fg)

        blend_label = tk.Label(master, text="Blend Mode:").place(x=5, y=50 + 65)

        blend_combo = ttk.Combobox(master, width=15, values=PsBlendMode)
        blend_combo.place(x=85, y=50 + 65)
        blend_combo.current(PsBlendMode.index(self.blend_mode))
        blend_combo.bind("<<ComboboxSelected>>", set_blend_mode)

        tk.Button(master, text="<prev background", width=20, command=back_bg).place(
            x=25, y=85 + 62
        )
        tk.Button(master, text="next background>", width=20, command=next_bg).place(
            x=175, y=85 + 62
        )

        tk.Button(master, text="<prev foreground", width=20, command=back_fg).place(
            x=25, y=125 + 62
        )
        tk.Button(master, text="next foreground>", width=20, command=next_fg).place(
            x=175, y=125 + 62
        )

        tk.Button(master, text="Save", width=20, command=save).place(x=400, y=100 + 65)

        master.bind("<Escape>", lambda e: exit())
        tk.mainloop()


tool = toolbox()
tool.run()
