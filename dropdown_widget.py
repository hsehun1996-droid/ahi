# -*- coding: utf-8 -*-
"""Shared button-style dropdown widgets for the app."""

import tkinter as tk

import customtkinter as ctk


def _get_master_bg_color(master):
    try:
        color = master.cget("fg_color")
        if isinstance(color, tuple):
            return color
        if color not in (None, "transparent"):
            return color
    except Exception:
        pass
    return "transparent"


class ButtonDropdown(ctk.CTkFrame):
    """A compact dropdown that mimics the app's button-style menu."""

    def __init__(
        self,
        master,
        width=140,
        height=30,
        corner_radius=10,
        border_width=1,
        border_color="#C8D7EA",
        fg_color="#FFFFFF",
        hover_color="#E9F1FA",
        text_color="#213547",
        font=None,
        dropdown_fg_color="#FFFFFF",
        dropdown_hover_color="#E8F1FB",
        dropdown_text_color="#213547",
        dropdown_font=None,
        values=None,
        variable=None,
        state="normal",
        command=None,
        anchor="w",
        button_color=None,
        button_hover_color=None,
        **kwargs,
    ):
        if button_color is not None:
            fg_color = button_color if fg_color == "#FFFFFF" else fg_color
        if button_hover_color is not None:
            hover_color = button_hover_color if hover_color == "#E9F1FA" else hover_color
        self._bg_color = _get_master_bg_color(master)
        super().__init__(
            master,
            width=width,
            height=height,
            fg_color="transparent",
            bg_color=self._bg_color,
            **kwargs,
        )
        self.grid_propagate(False)
        self.pack_propagate(False)

        self._width = width
        self._height = height
        self._corner_radius = corner_radius
        self._border_width = border_width
        self._border_color = border_color
        self._fg_color = fg_color
        self._hover_color = hover_color
        self._text_color = text_color
        self._font = font
        self._dropdown_fg_color = dropdown_fg_color
        self._dropdown_hover_color = dropdown_hover_color
        self._dropdown_text_color = dropdown_text_color
        self._dropdown_font = dropdown_font or font
        self._anchor = anchor
        self._command = command
        self._state = state
        self._values = [str(v) for v in (values or [])]
        self._menu = None

        if variable is None:
            initial_value = self._values[0] if self._values else ""
            self._variable = tk.StringVar(master=self, value=initial_value)
        else:
            self._variable = variable
            if not self._variable.get() and self._values:
                self._variable.set(self._values[0])

        self._button = ctk.CTkButton(
            self,
            text="",
            width=width,
            height=height,
            bg_color=self._bg_color,
            corner_radius=corner_radius,
            border_width=border_width,
            border_color=border_color,
            fg_color=fg_color,
            hover_color=hover_color,
            text_color=text_color,
            font=font,
            anchor=anchor,
            command=self._open_menu,
        )
        self._button.pack(fill="both", expand=True)

        self._variable_trace_id = self._variable.trace_add("write", self._on_variable_changed)
        self._refresh_display()
        self._apply_state()

    def _format_text(self, value):
        label = value if value else "선택"
        return f"{label}  ▾"

    def _refresh_display(self):
        self._button.configure(text=self._format_text(self._variable.get()))

    def _apply_state(self):
        if self._state == "disabled":
            self._button.configure(state="disabled")
        else:
            self._button.configure(state="normal")

    def _on_variable_changed(self, *_args):
        self._refresh_display()

    def _rebuild_menu(self):
        if self._menu is not None:
            self._menu.destroy()

        menu = tk.Menu(
            self,
            tearoff=0,
            bg=self._dropdown_fg_color,
            fg=self._dropdown_text_color,
            activebackground=self._dropdown_hover_color,
            activeforeground=self._dropdown_text_color,
            relief="flat",
            bd=1,
            activeborderwidth=0,
            font=self._dropdown_font,
        )
        for value in self._values:
            menu.add_command(label=value, command=lambda selected=value: self._select_value(selected))
        self._menu = menu

    def _open_menu(self):
        if self._state == "disabled":
            return
        self._rebuild_menu()
        self.update_idletasks()
        try:
            self._menu.tk_popup(self.winfo_rootx(), self.winfo_rooty() + self.winfo_height() + 4)
        finally:
            self._menu.grab_release()

    def _select_value(self, value):
        self._variable.set(value)
        if callable(self._command):
            self._command(value)

    def get(self):
        return self._variable.get()

    def set(self, value):
        value = "" if value is None else str(value)
        if value and value not in self._values:
            self._values.append(value)
        self._variable.set(value)

    def configure(self, require_redraw=False, cnf=None, **kwargs):
        if cnf:
            kwargs.update(cnf)

        frame_kwargs = {}
        button_kwargs = {}

        if "values" in kwargs:
            self._values = [str(v) for v in (kwargs.pop("values") or [])]
            current = self._variable.get()
            if self._values and current not in self._values:
                self._variable.set(self._values[0])
            elif not self._values:
                self._variable.set("")
        if "command" in kwargs:
            self._command = kwargs.pop("command")
        if "state" in kwargs:
            self._state = kwargs.pop("state")
        if "variable" in kwargs:
            new_var = kwargs.pop("variable")
            try:
                self._variable.trace_remove("write", self._variable_trace_id)
            except Exception:
                pass
            self._variable = new_var
            if not self._variable.get() and self._values:
                self._variable.set(self._values[0])
            self._variable_trace_id = self._variable.trace_add("write", self._on_variable_changed)
        if "width" in kwargs:
            self._width = kwargs["width"]
            frame_kwargs["width"] = kwargs["width"]
            button_kwargs["width"] = kwargs["width"]
        if "height" in kwargs:
            self._height = kwargs["height"]
            frame_kwargs["height"] = kwargs["height"]
            button_kwargs["height"] = kwargs["height"]
        if "corner_radius" in kwargs:
            self._corner_radius = kwargs["corner_radius"]
            button_kwargs["corner_radius"] = kwargs["corner_radius"]
        if "border_width" in kwargs:
            self._border_width = kwargs["border_width"]
            button_kwargs["border_width"] = kwargs["border_width"]
        if "border_color" in kwargs:
            self._border_color = kwargs["border_color"]
            button_kwargs["border_color"] = kwargs["border_color"]
        if "fg_color" in kwargs:
            self._fg_color = kwargs["fg_color"]
            button_kwargs["fg_color"] = kwargs["fg_color"]
        if "bg_color" in kwargs:
            self._bg_color = kwargs.pop("bg_color")
            frame_kwargs["bg_color"] = self._bg_color
            button_kwargs["bg_color"] = self._bg_color
        if "hover_color" in kwargs:
            self._hover_color = kwargs["hover_color"]
            button_kwargs["hover_color"] = kwargs["hover_color"]
        if "button_color" in kwargs:
            self._fg_color = kwargs.pop("button_color")
            button_kwargs["fg_color"] = self._fg_color
        if "button_hover_color" in kwargs:
            self._hover_color = kwargs.pop("button_hover_color")
            button_kwargs["hover_color"] = self._hover_color
        if "text_color" in kwargs:
            self._text_color = kwargs["text_color"]
            button_kwargs["text_color"] = kwargs["text_color"]
        if "font" in kwargs:
            self._font = kwargs["font"]
            button_kwargs["font"] = kwargs["font"]
        if "anchor" in kwargs:
            self._anchor = kwargs["anchor"]
            button_kwargs["anchor"] = kwargs["anchor"]
        if "dropdown_fg_color" in kwargs:
            self._dropdown_fg_color = kwargs.pop("dropdown_fg_color")
        if "dropdown_hover_color" in kwargs:
            self._dropdown_hover_color = kwargs.pop("dropdown_hover_color")
        if "dropdown_text_color" in kwargs:
            self._dropdown_text_color = kwargs.pop("dropdown_text_color")
        if "dropdown_font" in kwargs:
            self._dropdown_font = kwargs.pop("dropdown_font")

        super().configure(require_redraw=require_redraw, **frame_kwargs)
        if button_kwargs:
            self._button.configure(**button_kwargs)
        self._apply_state()
        self._refresh_display()

    config = configure

    def cget(self, attribute_name):
        if attribute_name == "values":
            return list(self._values)
        if attribute_name == "command":
            return self._command
        if attribute_name == "state":
            return self._state
        if attribute_name == "variable":
            return self._variable
        return super().cget(attribute_name)

    def destroy(self):
        if self._menu is not None:
            self._menu.destroy()
        try:
            self._variable.trace_remove("write", self._variable_trace_id)
        except Exception:
            pass
        super().destroy()
