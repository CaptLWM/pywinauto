from pywinauto import Application, Desktop

# SCADA 창을 찾을 때 사용할 키워드
# 창 제목에 이 문자열이 포함된 프로세스를 자동으로 찾는다.
TARGET_TITLE_KEYWORD = "SCADA"


def find_scada_pid():
    """SCADA 메인창을 찾아서 해당 프로세스 ID(PID)를 반환"""
    # UIA 기반 Desktop의 모든 top-level window를 탐색
    for w in Desktop(backend="uia").windows():
        # 창 제목을 소문자로 변환하여 키워드 검색
        if TARGET_TITLE_KEYWORD.lower() in (w.window_text() or "").lower():
            # 해당 창의 프로세스 ID 반환
            return w.process_id()
    return None  # 못 찾았으면 None


def dump_ui(ctrl, indent=0):
    """주어진 컨트롤(ctrl)의 UI 구조를 재귀적으로 출력"""
    prefix = "  " * indent  # 들여쓰기(트리 계층 구조 표현)

    # friendly_class_name(): 버튼, 메뉴, Static 같은 UI 클래스명 표시
    # window_text(): 해당 UI 요소의 텍스트(표시 문자열)
    print(f"{prefix}- [{ctrl.friendly_class_name()}] title='{ctrl.window_text()}'")

    # 자식 요소 출력
    try:
        for child in ctrl.children():   # 모든 하위 UI 요소 검색
            dump_ui(child, indent + 1)  # 한 단계 더 들여쓰기하여 재귀 호출
    except Exception:
        # 일부 컨트롤은 children() 호출이 안될 수도 있어 예외 무시
        pass


def main():
    """SCADA 프로그램의 모든 창과 UI 구조를 덤프하는 메인 함수"""

    # 1) SCADA 프로세스의 PID 찾기
    pid = find_scada_pid()
    if pid is None:
        print("SCADA PID 찾기 실패")
        return

    # 2) pywinauto로 해당 PID의 프로세스 연결
    # backend=uia를 사용하면 Windows 10/11의 대부분 UI를 지원
    app = Application(backend="uia").connect(process=pid)

    # 3) Desktop 객체 생성 (이걸 통해 모든 창을 스캔)
    desktop = Desktop(backend="uia")

    # 4) SCADA 프로세스가 가진 모든 top-level window 수집
    #    메인창, 설정창, 숨겨진 Dialog 등 포함
    all_windows = desktop.windows(process=pid)

    print(f"총 발견된 창 수: {len(all_windows)}\n")

    # 5) 각 창에 대해 UI 구조를 출력
    for idx, win in enumerate(all_windows):
        print(f"\n===== WINDOW #{idx+1}: {win.window_text()} =====")

        try:
            dump_ui(win)  # UI 트리 재귀 출력
        except Exception:
            print("[!] 덤프 중 오류 발생")


if __name__ == "__main__":
    main()  # 프로그램 실행
