from pywinauto import findwindows
from pywinauto import Application
import time

class CalculatorAutomation:
    # class 에서 사용할 변수 저장.
    def __init__(self):
        self.app = None
        self.dlg = None
        self.values = ["공"] + list(map(str, range(1, 10)))

    # 계산기 연결
    def connect_application(self):
        procs = findwindows.find_elements(top_level_only=True)
        for num, proc in enumerate(procs):
        # 계산기가 켜져있으면, 해당 애플리케이션으로 연결함.
            if "계산기" in str(proc):
                print(f"{num} : {proc} / 핸들 : {proc.handle} / 프로세스 : {proc.process_id}")
                # 만약 실행되고 있다면, 프로그램을 연결한다.
                # title, class_name, process, handle 등으로 구분하여 작성할 수 있다.
                try:
                    proc_value = proc.process_id
                    self.app = Application(backend='uia').connect(process = proc_value)
                    print(f"Found UIA Application : {proc}")
                    return
                except Exception as e:
                    # app = Application('uia').start("calc.exe")
                    print(f"Open a UIA Application {e}")

        self.app = Application().start("calc.exe")
        time.sleep(1) # 대기 신가 확보
        self.connect_application()

    # 숫자를 누르는 함수
    def press_number(self, num):
        # num을 문자열로 변환 후 각 자리를 int 형식으로 변환
        num_list = list(map(int, str(num)))
        # 각 버튼을 클릭함.
        for nl in num_list:
            self.click_button(self.values[nl])

    # 시작 숫자부터 마지막 숫자까지 더함
    def plus_range(self):
        start_num, end_num = map(int, input("시작 숫자와 종료 숫자를 입력하세요 (예: 1 100): ").split())
        
        # 일정 범위의 숫자를 덧셈하기.
        for i in range(start_num, end_num+1):
            self.press_number(i)
            self.click_button("더하기")
        result = self.dlg.child_window(control_type="Text", found_index = 2).window_text()
        print(f"결과값 : {result.replace("표시는 ", "")}")

    # 버튼이 보일 때까지 기다렸다가 클릭한다.
    def click_button(self, value):
        button = self.dlg.child_window(title=value, control_type="Button")
        button.wait('visible', timeout = 10)
        button.click()

    # 깨끗하게 지움
    def clear(self):
        # C버튼 클릭
        self.click_button("지우기")

    #  메인 함수.
    def main(self):
        self.connect_application()

        # 내부의 win 값을 전체 가져오기.
        for win in self.app.windows():
            print(f"Title: {win.window_text()}, Handle: {win.handle}, Process ID : {win.process_id()}, Class Name: {win.class_name()}")


        # 윈도우 창을 연결하기 위해 .window() 메서드를 사용함.
        # title, class_name 등을 사용할 수 있다.
        # dlg = app.window(process = )
        self.dlg = self.app.window(title="계산기")
        self.dlg.print_control_identifiers()
        
        # 공학용 계산기를 열기
        self.click_button("탐색 열기")
        self.dlg.child_window(title="공학용 계산기", auto_id="Scientific", control_type = "ListItem").select()

        # 표준 계산기로 돌아오기
        self.click_button("탐색 열기")
        self.dlg.child_window(title="표준 계산기", auto_id="Standard", control_type = "ListItem").select()

        n = 0
        while(n >= 0):
            function_list = [self.plus_range]
            for i, fl in enumerate(function_list):
                print(f"{i}. {fl.__name__}")
            n = int(input(f"0 부터 {len(function_list) - 1 } 사이의 값을 입력하세요 : "))
            # 1~21 더하기 수행
            function_list[n]()

if __name__ == "__main__":
    automation = CalculatorAutomation()
    automation.main()