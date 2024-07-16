import tkinter as tk
from PIL import Image, ImageTk

class location_btn():
    def __init__(self, button, id):
        self.button = button
        self.id = id

    def set_config(self, alias, label, func):
        self.button.config(command=lambda: func(alias, label, self.id))