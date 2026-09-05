"""Allow tkinter analysis when the build Python cannot initialize a display."""


def pre_find_module_path(api):
    # The matching hook-_tkinter.py collects the runtime explicitly.
    return None
