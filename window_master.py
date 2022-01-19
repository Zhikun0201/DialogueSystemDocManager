try:
    import unreal as ue
except ModuleNotFoundError:
    ue = None

def show_window(win):
    win.show()
    if ue:
        ue.parent_external_window_to_slate(win.winId())
    else:
        pass


def batch_lst_resize_mode(arr, typ):
    for a in arr:
        c = a[0]
        for k in range(1, len(a)):
            c.horizontalHeader().setSctionResizeMode(a[k], typ)


def window_exists(w):
    if w is None:
        return False


def window_delete(w):
    if w is None:
        return
    try:
        w.close()
        w.deleteLater()
    except (NameError, AttributeError, RuntimeError):
        pass
