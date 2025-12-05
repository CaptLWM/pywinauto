from pywinauto import Application, Desktop
import time
import traceback

OUTFILE = "ui_dump.txt"
ACTION = "NO_SAVE"   # SAVE / NO_SAVE / CANCEL
WAIT_POPUP = 1.0


def dump_obj(obj, indent=0, out_lines=None):
    if out_lines is None:
        out_lines = []
    try:
        text = obj.window_text()
    except Exception:
        text = ""
    try:
        ctrl_type = obj.friendly_class_name()
    except Exception:
        ctrl_type = str(type(obj))
    line = "  " * indent + f"- [{ctrl_type}] title='{text}'"
    out_lines.append(line)
    try:
        children = obj.children()
    except Exception:
        children = []
    for c in children:
        dump_obj(c, indent + 1, out_lines)
    return out_lines


def find_modal_in_parent(parent):
    candidates = []
    try:
        for c in parent.children():
            try:
                typ = c.element_info.control_type
            except Exception:
                typ = None
            if typ in ("Window", "Pane") or "Dialog" in str(typ):
                candidates.append(c)
    except Exception:
        pass

    for c in candidates:
        try:
            btns = [b for b in c.children() if b.element_info.control_type == "Button"]
            if btns:
                return c
        except Exception:
            continue

    if candidates:
        return candidates[0]
    return None


def click_button_by_patterns(window, patterns):
    for p in patterns:
        try:
            btn = window.child_window(title_re=p, control_type="Button")
            if btn.exists(timeout=0.5):
                btn.click_input()
                return True, p
        except Exception:
            continue
    return False, None


def main():
    try:
        app = Application(backend="uia").start("notepad.exe")
    except Exception:
        app = None

    desktop = Desktop(backend="uia")

    try:
        dlg = desktop.window(title_re=".*메모장.*|.*Notepad.*")
        dlg.wait("ready", timeout=5)
    except Exception:
        print("메모장 창을 찾지 못했습니다.")
        return

    dlg.set_focus()
    time.sleep(0.3)

    try:
        editor = dlg.child_window(control_type="Document")
        if editor.exists(timeout=1):
            editor.type_keys("Hello, world!", with_spaces=True)
    except Exception:
        pass

    dlg.type_keys("%{F4}")
    time.sleep(WAIT_POPUP)

    out_lines = ["--- Main window control tree ---"]
    out_lines += dump_obj(dlg)

    modal = find_modal_in_parent(dlg)
    if modal:
        out_lines.append("\n--- Modal detected ---")
        out_lines += dump_obj(modal, indent=1)
    else:
        out_lines.append("\n--- No modal detected ---")

    with open(OUTFILE, "w", encoding="utf-8") as f:
        f.write("\n".join(out_lines))

    print(f"UI 덤프 -> {OUTFILE}")

    if not modal:
        print("❌ 모달을 찾지 못함. ui_dump.txt 붙여줘.")
        return

    if ACTION == "SAVE":
        patterns = [r".*저장.*", r".*Save.*"]
    elif ACTION == "NO_SAVE":
        patterns = [r".*저장하지 않음.*", r".*저장 안 함.*", r".*Don't Save.*"]
    else:
        patterns = [r".*취소.*", r".*Cancel.*"]

    ok, pat = click_button_by_patterns(modal, patterns)
    print("클릭:", ok, "패턴:", pat)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("스크립트 실행 중 예외 발생:", e)
        traceback.print_exc()
