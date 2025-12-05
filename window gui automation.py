# Windows 11 Notepad 안정화 자동화 — Desktop 기반 모달 탐지 (덤프 기반 안정화)
from pywinauto import Application, Desktop
from pywinauto.keyboard import send_keys
import time
import tkinter as tk

# 동작 선택: "SAVE" / "NO_SAVE" / "CANCEL"
ACTION = "SAVE"

# --- clipboard helper using tkinter ---
def set_clipboard_text(text: str):
    root = tk.Tk()
    root.withdraw()
    root.clipboard_clear()
    root.clipboard_append(text)
    root.update()
    root.destroy()

# --- 버튼 클릭 헬퍼 ---
def click_button_by_patterns(window, patterns):
    for p in patterns:
        try:
            btn = window.child_window(title_re=p, control_type="Button")
            if btn.exists(timeout=0.6):
                btn.click_input()
                return True, p
        except Exception:
            continue
    # fallback: brute force scan 버튼 텍스트 포함 검사
    try:
        for b in window.children():
            try:
                if b.friendly_class_name() == "Button":
                    txt = (b.window_text() or "").lower()
                    for pat in ["저장", "저장 안", "저장하지", "don't save", "cancel", "취소"]:
                        if pat in txt:
                            b.click_input()
                            return True, txt
            except Exception:
                continue
    except Exception:
        pass
    return False, None

# --- 모달 찾기 (Desktop 전역 탐색, 덤프 구조 반영) ---
def find_modal_desktop(desktop):
    """
    덤프 기반 안정적 탐지:
    1) top-level 혹은 데스크탑에서 제목이 정확히 '메모장'인 윈도우 검색
    2) 없으면 모든 윈도우(descendants 포함) 중 friendly_class_name == 'Dialog' 이고 Button 자식이 있는 요소 반환
    """
    # 1) 직접 타이틀로 찾기 (덤프에 '메모장' 내부 Dialog 존재)
    try:
        w = desktop.window(title="메모장")
        if w.exists(timeout=0.6):
            return w
    except Exception:
        pass

    # 2) 타이틀 정규식으로 찾기
    try:
        w = desktop.window(title_re=".*메모장.*|.*Save.*")
        if w.exists(timeout=0.6):
            return w
    except Exception:
        pass

    # 3) 모든 최상위 윈도우 검색 후, descendant 중 Dialog-like 요소 찾기
    try:
        for top in desktop.windows():
            try:
                # 직접 top에 버튼이 있으면 후보
                btns = [c for c in top.children() if c.friendly_class_name() == "Button"]
                if btns:
                    # 만약 타이틀/텍스트에 '저장' 관련 문구가 있으면 채택
                    t = (top.window_text() or "").lower()
                    if "저장" in t or "save" in t or "변경 내용을" in t:
                        return top
                # descendant 내부에서 Dialog-like 요소 찾기
                for desc in top.descendants():
                    try:
                        if desc.friendly_class_name() == "Dialog":
                            # Dialog 내부에 버튼이 있는지 확인
                            try:
                                btns2 = [b for b in desc.children() if b.friendly_class_name() == "Button"]
                            except Exception:
                                btns2 = []
                            if btns2:
                                return desc
                    except Exception:
                        continue
            except Exception:
                continue
    except Exception:
        pass

    return None

# --- 메인 로직 ---
def main():
    desktop = Desktop(backend="uia")
    try:
        Application(backend="uia").start("notepad.exe")
    except Exception:
        pass

    # Notepad 창 찾기
    dlg = None
    try:
        dlg = desktop.window(title_re=".*메모장.*|.*Notepad.*")
        dlg.wait("ready", timeout=5)
    except Exception:
        for w in desktop.windows():
            t = w.window_text() or ""
            if "메모장" in t or "notepad" in t.lower():
                dlg = w
                break

    if dlg is None:
        print("메모장 창을 찾지 못했습니다.")
        return

    # 안전한 입력: 클립보드 붙여넣기
    text_to_write = "Hello, world!"
    set_clipboard_text(text_to_write)

    dlg.set_focus()
    time.sleep(0.2)

    try:
        editor = dlg.child_window(control_type="Document")
        if editor.exists(timeout=1):
            editor.set_focus()
            time.sleep(0.1)
            send_keys("^v")  # Ctrl+V
            # 붙여넣기 반영 확인
            for _ in range(10):
                cur = (editor.window_text() or "")
                if text_to_write in cur:
                    break
                time.sleep(0.15)
    except Exception:
        dlg.set_focus()
        time.sleep(0.1)
        send_keys("^v")
        time.sleep(0.2)

    # 닫기 및 모달 탐지/처리
    attempts = 0
    while attempts < 2:
        attempts += 1
        try:
            dlg.close()
        except Exception:
            dlg.type_keys("%{F4}")
        time.sleep(0.7)

        modal = find_modal_desktop(desktop)
        if modal:
            print("모달 감지:", modal.window_text())
            modal.set_focus()
            time.sleep(0.2)

            if ACTION == "SAVE":
                patterns = [r".*저장$", r".*Save"]
            elif ACTION == "NO_SAVE":
                patterns = [r"저장\s*하지\s*않음", r"저장\s*안\s*함", r".*Don't\s*Save.*", r".*Don't\s*Save"]
            else:
                patterns = [r".*취소.*", r".*Cancel.*"]

            ok, pat = click_button_by_patterns(modal, patterns)
            if ok:
                print("버튼 클릭 성공:", pat)
                return
            else:
                print("내부 모달에서 버튼을 못찾음 — 대체 전략 시도")

                # 대체 전략: modal의 모든 버튼 출력(디버그) 후 브루트 포스 클릭 시도
                try:
                    for b in modal.children():
                        try:
                            if b.friendly_class_name() == "Button":
                                print("버튼 발견:", b.window_text())
                                txt = (b.window_text() or "").lower()
                                if any(x in txt for x in ["저장", "save", "취소", "don't"]):
                                    b.click_input()
                                    print("브루트포스 클릭:", b.window_text())
                                    return
                        except Exception:
                            continue
                except Exception:
                    pass

                print("모달에서 버튼 클릭 실패")
                return

        # Save As (top-level) 확인
        saveas = None
        try:
            saveas = find_modal_desktop(desktop)  # 이미 시도했으므로 같은 함수 재사용 (안정성)
        except Exception:
            saveas = None

        if saveas and saveas is not modal:
            print("Top-level SaveAs 감지:", saveas.window_text())
            try:
                ok, pat = click_button_by_patterns(saveas, [r".*취소.*", r".*Cancel.*"])
                if ok:
                    print("Save As에서 취소 클릭 -> 재시도")
                    time.sleep(0.3)
                    continue
            except Exception:
                pass

        print("팝업 없음. 시도 횟수:", attempts)
        break

    print("자동화 종료.")

if __name__ == "__main__":
    main()
