# %%
import tkinter as tk
from tkinter import ttk
from pyhwpx import Hwp

# region 초기화 부분
# 인스턴스생성
hwp = Hwp()

# Tkinter 창 생성
root = tk.Tk()
root.title("송나겸 돈 복사 매크로(v1.0)")

# 파일 경로 표시
file_path_frame = tk.Frame(
    root, highlightbackground="black", highlightthickness=1, padx=10, pady=5
)
file_path_frame.pack(fill="x", padx=10, pady=5)
# 공지사항 라벨
notice_label = tk.Label(
    file_path_frame,
    text="2002방식 조판부호 사용 해야 작동함. / 창이 활성화된 상태에서만 단축키가 작동함.",
    font=("Arial", 10),
    anchor="w",
)
notice_label.grid(
    row=1, column=0, columnspan=2, sticky="w", padx=5
)  # 아래쪽 전체 너비로 표시

# 파일 경로 라벨
file_path_label = tk.Label(
    file_path_frame, text="파일 경로: 없음", font=("Arial", 10), anchor="w"
)
file_path_label.grid(row=0, column=0, sticky="w", padx=5)  # 왼쪽 정렬

hwp.Run("FileOpen")


def update_file_path_label():
    """파일 경로를 업데이트하고 Label에 표시"""
    if hwp.Path:
        file_path = hwp.Path
        file_path_label.config(text=f"파일 경로: {file_path}")
        print(f"파일 경로: {file_path}")
    else:
        file_path_label.config(text="파일 경로: 없음")
        print("파일 경로: 없음")


def select_file():
    """파일을 재선택"""
    try:
        hwp.Save()
        hwp.Run("FileClose")
        hwp.Run("FileOpen")
        update_file_path_label()
    except Exception as e:
        print(f"파일 재선택 중 오류: {e}")


# 파일 재선택 버튼
file_reselect_button = tk.Button(
    file_path_frame, text="파일 재선택", command=select_file
)
file_reselect_button.grid(row=0, column=1, sticky="e", padx=5)  # 오른쪽 정렬

# 표 초기화
Table_list = []  # 모든 표의 갯수 찾기
for i in hwp.ctrl_list:
    if i.UserDesc == "표":  # 컨트롤이 표일 경우
        Table_list.append(i)  # 리스트에 저장
Table_Total = len(Table_list)  # 문서안의 모든 표의 갯수
Table_index = 0  # 전역으로 사용할 표 인덱스

# Notebook(탭 컨트롤러) 생성
notebook = ttk.Notebook(root)
notebook.pack(expand=True, fill="both", padx=10, pady=10)
# endregion


# region 매크로 함수
def 셀_전체선택():
    hwp.TableCellBlock()
    hwp.TableCellBlockExtend()
    hwp.TableCellBlockExtend()


def 셀_전체선택ex():
    hwp.ShapeObjTableSelCell()
    hwp.TableCellBlockExtend()
    hwp.TableCellBlockExtend()


def 표라인_전체투명():
    hwp.TableCellBorderNo()


def 표라인_안쪽실선():
    hwp.TableCellBorderInside()


def 표라인_위아래굵은선():
    pset = hwp.HParameterSet.HCellBorderFill
    hwp.HAction.GetDefault("CellBorder", pset.HSet)
    pset.BorderWidthBottom = hwp.HwpLineWidth("0.4mm")
    pset.BorderTypeBottom = hwp.HwpLineType("Solid")
    pset.BorderWidthTop = hwp.HwpLineWidth("0.4mm")
    pset.BorderTypeTop = hwp.HwpLineType("Solid")
    return hwp.HAction.Execute("CellBorder", pset.HSet)


def 표라인_양옆투명():
    pset = hwp.HParameterSet.HCellBorderFill
    hwp.HAction.GetDefault("CellBorder", pset.HSet)
    pset.BorderTypeRight = hwp.HwpLineType("None")
    pset.BorderTypeLeft = hwp.HwpLineType("None")
    pset.FillAttr.GradationAlpha = 0
    pset.FillAttr.ImageAlpha = 0
    hwp.HAction.Execute("CellBorder", pset.HSet)


def 표색_없음():  # 안됨
    # hwp.cell_fill([255,255,255]) #흰색으로 채우기
    pset = hwp.HParameterSet.HCellBorderFill
    hwp.HAction.GetDefault("CellBorder", pset.HSet)
    pset.FillAttr.type = hwp.BrushType("NullBrush")  # 소문자임.. ㅜㅜㅜㅜㅜ
    pset.FillAttr.GradationAlpha = 0
    pset.FillAttr.WindowsBrush = 0
    pset.FillAttr.ImageAlpha = 0
    return hwp.HAction.Execute("CellBorder", pset.HSet)  # false,,,,,,,


def 셀세로정렬():
    pset = hwp.HParameterSet.HShapeObject
    hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
    pset.HSet.SetItem("ShapeType", 3)
    pset.HSet.SetItem("ShapeCellSize", 0)
    pset.ShapeTableCell.VertAlign = hwp.VAlign("Center")
    hwp.HAction.Execute("TablePropertyDialog", pset.HSet)


def 안여백지정해제():
    pset = hwp.HParameterSet.HShapeObject
    hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
    pset.HSet.SetItem("ShapeType", 3)
    pset.HSet.SetItem("ShapeCellSize", 0)
    pset.ShapeTableCell.HasMargin = 0
    hwp.HAction.Execute("TablePropertyDialog", pset.HSet)


def 글자처럼해제_자리차지():
    pset = hwp.HParameterSet.HShapeObject
    hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
    pset.TextWrap = hwp.TextWrapType("TopAndBottom")
    pset.TreatAsChar = 0
    pset.HSet.SetItem("ShapeType", 3)
    pset.HSet.SetItem("ShapeCellSize", 0)
    hwp.HAction.Execute("TablePropertyDialog", pset.HSet)


def 셀단위로나눔_제목줄반복():
    pset = hwp.HParameterSet.HShapeObject
    hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
    pset.RepeatHeader = 1
    pset.PageBreak = hwp.TableBreak("Table")
    pset.HSet.SetItem("ShapeType", 3)
    pset.HSet.SetItem("ShapeCellSize", 0)
    hwp.HAction.Execute("TablePropertyDialog", pset.HSet)


def 표색_헤드_회색217():
    hwp.TableColPageUp()
    hwp.cell_fill([217, 217, 217])


def 표라인_헤드_두줄():
    hwp.TableColPageUp()
    pset = hwp.HParameterSet.HCellBorderFill
    hwp.HAction.GetDefault("CellBorder", pset.HSet)
    pset.BorderWidthBottom = hwp.HwpLineWidth("0.5mm")
    pset.BorderTypeBottom = hwp.HwpLineType("DoubleSlim")
    hwp.HAction.Execute("CellBorder", pset.HSet)


def 표좌우맞춤():
    pagedef_dict = hwp.get_pagedef_as_dict()
    paper_width = int(pagedef_dict["용지폭"])  # 용지 폭
    margin_left = int(pagedef_dict["왼쪽"])  # 왼쪽 여백
    margin_right = int(pagedef_dict["오른쪽"])  # 오른쪽 여백
    # 실제 문서 폭 계산
    actual_width = paper_width - margin_left - margin_right
    epsilon = align_ep.get()
    hwp.set_table_width(actual_width - epsilon)
    셀_전체선택()


def 위캡션2mm():
    hwp.ShapeObjAttachCaption()
    pset = hwp.HParameterSet.HShapeObject
    hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
    pset.HSet.SetItem("ShapeType", 3)
    pset.ShapeCaption.Side = hwp.SideType("Top")
    pset.ShapeCaption.Width = hwp.MiliToHwpUnit(30.0)
    pset.ShapeCaption.Gap = hwp.MiliToHwpUnit(2.0)
    pset.ShapeCaption.CapFullSize = 0
    pset.HSet.SetItem("ShapeCellSize", 0)
    hwp.HAction.Execute("TablePropertyDialog", pset.HSet)
    # 표 안쪽으로 들어가는 코드 추가
    hwp.MoveDown()


def 아래캡션3mm():
    hwp.ShapeObjAttachCaption()
    pset = hwp.HParameterSet.HShapeObject
    hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
    pset.HSet.SetItem("ShapeType", 3)
    pset.ShapeCaption.Side = hwp.SideType("Bottom")
    pset.ShapeCaption.Width = hwp.MiliToHwpUnit(30.0)
    pset.ShapeCaption.Gap = hwp.MiliToHwpUnit(3.0)
    pset.ShapeCaption.CapFullSize = 0
    pset.HSet.SetItem("ShapeCellSize", 0)
    hwp.HAction.Execute("TablePropertyDialog", pset.HSet)
    # 표 안쪽으로 들어가는 코드 추가
    hwp.MoveUp()


def 표위치_2단(Horz: str, Vert: str):
    """Horz은  Center(왼),Left(가운데),Right(오른)  /
    Vert는 Top, Bottom
    """
    pset = hwp.HParameterSet.HShapeObject
    hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
    pset.HorzAlign = hwp.HAlign(Horz)
    pset.HorzRelTo = hwp.HorzRel("Page")
    pset.VertAlign = hwp.VAlign(Vert)
    pset.VertRelTo = hwp.VertRel("Page")
    if Vert == "Top":
        pset.OutsideMarginBottom = hwp.MiliToHwpUnit(7.0)
    else:
        pset.OutsideMarginBottom = hwp.MiliToHwpUnit(0.0)
    if Vert == "Bottom":
        pset.OutsideMarginTop = hwp.MiliToHwpUnit(7.0)
    else:
        pset.OutsideMarginTop = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginRight = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginLeft = hwp.MiliToHwpUnit(0.0)
    pset.HSet.SetItem("ShapeType", 3)
    pset.HSet.SetItem("ShapeCellSize", 0)
    hwp.HAction.Execute("TablePropertyDialog", pset.HSet)


def 표위치_1단():
    """Horz은  Center(왼),Left(가운데),Right(오른)  /
    Vert는 Top, Bottom
    """
    pset = hwp.HParameterSet.HShapeObject
    hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
    pset.HorzAlign = hwp.HAlign("Left")
    pset.HorzRelTo = hwp.HorzRel("Page")
    pset.VertAlign = hwp.VAlign("Top")
    pset.VertRelTo = hwp.VertRel("Para")
    pset.OutsideMarginBottom = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginTop = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginRight = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginLeft = hwp.MiliToHwpUnit(0.0)
    pset.HSet.SetItem("ShapeType", 3)
    pset.HSet.SetItem("ShapeCellSize", 0)
    hwp.HAction.Execute("TablePropertyDialog", pset.HSet)


# endregion


# region 함수 모음
def on_CellLineMacro(event=None):
    """버튼 클릭 이벤트: 드롭다운 값 출력"""
    first_option = dropdown1.get()
    second_option = dropdown2.get()
    print(f"첫행 배경색: {first_option}, 첫행 밑 줄: {second_option}")
    # 표라인_전체투명()
    # 표라인_안쪽실선()
    표라인_위아래굵은선()
    표라인_양옆투명()
    표색_없음()
    hwp.set_table_inside_margin(1, 1, 1, 1)  # 안여벡 1mm로 밀기
    hwp.TableVAlignCenter()  # 셀 세로 중앙정렬
    안여백지정해제()
    글자처럼해제_자리차지()  # 1 위계가 있음.
    셀단위로나눔_제목줄반복()  # 2 위계가 있음.
    위캡션2mm()
    셀_전체선택()
    hwp.set_table_outside_margin(0, 0, 0, 0)  # 밖여백 0mm로 밀기
    셀_전체선택()
    if first_option == "없음":
        if second_option == "한줄":
            print("옵션: 없음, 한줄 선택")
        elif second_option == "두줄":
            print("옵션: 없음, 두줄 선택")
            표라인_헤드_두줄()
    elif first_option == "회색 255,255,255,15%":
        표색_헤드_회색217()
        셀_전체선택()
        if second_option == "한줄":
            print("옵션: 회색, 한줄 선택")
        elif second_option == "두줄":
            print("옵션: 회색, 두줄 선택")
            표라인_헤드_두줄()
    else:
        print("선택되지 않은 옵션이 있습니다.")


def 양옆맞추기(event=None):
    """양옆맞추기(오차0.4임)"""
    표좌우맞춤()
    # 셀_전체선택ex()


def 단2맞추기(event=None):
    """2단일때 양옆 맞추기"""
    print("2단맞추기")


def 그림용(event=None):
    """그림 표용 세팅"""
    표라인_전체투명()
    표색_없음()
    hwp.set_table_inside_margin(1, 1, 1, 1)  # 안여벡 1mm로 밀기
    hwp.TableVAlignCenter()  # 셀 세로 중앙정렬
    안여백지정해제()
    글자처럼해제_자리차지()  # 1 위계가 있음.
    셀단위로나눔_제목줄반복()  # 2 위계가 있음.
    아래캡션3mm()
    셀_전체선택()
    hwp.set_table_outside_margin(0, 0, 0, 0)  # 밖여백 0mm로 밀기
    셀_전체선택()


def handle_1단용(event=None):
    표위치_1단()


def handle_2단_왼_상(event=None):
    표위치_2단("Center", "Top")


def handle_2단_가운_상(event=None):
    표위치_2단("Left", "Top")


def handle_2단_오른_상(event=None):
    표위치_2단("Right", "Top")


def handle_2단_왼_하(event=None):
    표위치_2단("Center", "Bottom")


def handle_2단_가운_하(event=None):
    표위치_2단("Left", "Bottom")


def handle_2단_오른_하(event=None):
    표위치_2단("Right", "Bottom")


def 이전표(event=None):
    global Table_index
    global Table_list
    # Spinbox에서 현재 값을 가져오고 -1
    current_value = current_table_index.get()
    prev_index = max(0, current_value - 1)  # 최소값 제한
    current_table_index.set(prev_index)  # Spinbox에 값 설정
    Table_index = prev_index  # 전역 변수 업데이트
    print(f"현재 Table_index: {Table_index}")
    hwp.get_into_nth_table(Table_index)  # 인덱스 표의 첫번째 셀로 이동
    셀_전체선택()


def 다음표(event=None):
    global Table_index
    global Table_list
    # Spinbox에서 현재 값을 가져오고 +1
    current_value = current_table_index.get()
    next_index = max(0, min(current_value + 1, Table_Total - 1))  # 최대값 제한
    current_table_index.set(next_index)  # Spinbox에 값 설정
    Table_index = next_index  # 전역 변수 업데이트
    print(f"현재 Table_index: {Table_index}")
    hwp.get_into_nth_table(Table_index)  # 인덱스 표의 첫번째 셀로 이동
    셀_전체선택()


def 처음으로(event=None):
    """처음으로"""
    global Table_index
    global Table_list
    # Spinbox와 Table_index를 동기화
    Table_index = 0  # 전역 변수 Table_index를 0으로 설정
    current_table_index.set(Table_index)  # Spinbox에 값 반영
    try:
        hwp.get_into_nth_table(Table_index)  # 첫 번째 표로 이동
        셀_전체선택()
    except Exception as e:
        print(f"처음으로 이동 중 오류 발생: {e}")


# endregion

# region 첫번째 탭 : 표 셀/선 매크로
##########################################
tab1 = ttk.Frame(notebook)
notebook.add(tab1, text="표 셀/선 매크로")


# 키 이벤트 연결
def bind_button_to_key(button, key):
    """버튼을 특정 키에 바인딩"""

    def key_action(event=None):
        # 버튼 눌린 모션 추가
        if event.keysym.lower() == key.lower():
            # 버튼 눌린 모션 추가
            button.config(relief="sunken")  # 눌림 효과
            button.update_idletasks()  # 즉시 업데이트
            button.invoke()  # 버튼의 command 실행
            root.after(
                10, lambda: button.config(relief="raised")
            )  # 0.01초 후 원래 상태로 복구

    # 모든 키 입력을 감지하여 처리
    root.bind_all(f"<KeyPress-{key.lower()}>", key_action)
    root.bind_all(f"<KeyPress-{key.upper()}>", key_action)


# '표 이동' 프레임
frame_table_movement = tk.Frame(
    tab1, highlightbackground="black", highlightthickness=2, padx=10, pady=10
)
frame_table_movement.grid(row=0, column=0, padx=30, pady=10)

position_description = tk.Label(
    frame_table_movement, text="표 이동", font=("Arial", 12, "bold")
)
position_description.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

align_ep_label = tk.Label(frame_table_movement, text=f"전체 표 갯수={Table_Total}")
align_ep_label.grid(row=0, column=2, padx=5, pady=5, sticky="w")

current_table_index = tk.IntVar(value=0)
align_ep = tk.DoubleVar(value=0.4)

# 현재 표 인덱스와 이동 버튼
current_index_label = tk.Label(frame_table_movement, text="현재 표 인덱스:")
current_index_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")

index_spinbox = tk.Spinbox(
    frame_table_movement, from_=0, to=0, textvariable=current_table_index, width=5
)
index_spinbox.grid(row=1, column=1, padx=5, pady=5)

btn_prev = tk.Button(frame_table_movement, text="이전[W]", command=이전표)
btn_prev.grid(row=2, column=1, padx=5, pady=5)
bind_button_to_key(btn_prev, "w")

btn_next = tk.Button(frame_table_movement, text="다음[S]", command=다음표)
btn_next.grid(row=3, column=1, padx=5, pady=5)
bind_button_to_key(btn_next, "s")

btn_begin = tk.Button(frame_table_movement, text="처음으로[B]", command=처음으로)
btn_begin.grid(row=1, column=2, columnspan=3, padx=5, pady=5)
bind_button_to_key(btn_begin, "b")

처음으로()
# '배경/선 매크로' 프레임
frame_cell_macro = tk.Frame(
    tab1, highlightbackground="black", highlightthickness=2, padx=10, pady=10
)
frame_cell_macro.grid(row=1, column=0, padx=30, pady=10)

cell_macro_label = tk.Label(
    frame_cell_macro, text="배경/선/여백/정렬 매크로", font=("Arial", 12, "bold")
)
cell_macro_label.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

dropdown1_label = tk.Label(frame_cell_macro, text="첫행 배경색:")
dropdown1_label.grid(row=1, column=0, padx=5, pady=5)

dropdown1 = ttk.Combobox(frame_cell_macro, values=["없음", "회색 255,255,255,15%"])
dropdown1.set("회색 255,255,255,15%")
dropdown1.grid(row=1, column=1, padx=5, pady=5)

dropdown2_label = tk.Label(frame_cell_macro, text="첫행 밑 줄:")
dropdown2_label.grid(row=2, column=0, padx=5, pady=5)

dropdown2 = ttk.Combobox(frame_cell_macro, values=["한줄", "두줄"])
dropdown2.set("한줄")
dropdown2.grid(row=2, column=1, padx=5, pady=5)

btn_execute = tk.Button(frame_cell_macro, text="실행[D]", command=on_CellLineMacro)
btn_execute.grid(row=3, column=0, columnspan=2, padx=5, pady=5)
bind_button_to_key(btn_execute, "d")

btn_align = tk.Button(frame_cell_macro, text="전체 좌우 맞추기[F]", command=양옆맞추기)
btn_align.grid(row=1, column=2, columnspan=2, padx=5, pady=5)
bind_button_to_key(btn_align, "f")

align_ep_label = tk.Label(frame_cell_macro, text="e = ")
align_ep_label.grid(row=0, column=3, padx=5, pady=5, sticky="w")

index_spinbox = tk.Spinbox(
    frame_cell_macro, from_=0, to=1, increment=0.1, textvariable=align_ep, width=3
)
index_spinbox.grid(row=0, column=3, padx=5, pady=5)

btn_align2 = tk.Button(
    frame_cell_macro, text="2단 좌우 맞추기(미완)[G]", command=단2맞추기
)
btn_align2.grid(row=2, column=2, columnspan=2, padx=5, pady=5)
bind_button_to_key(btn_align2, "g")

btn_pic = tk.Button(frame_cell_macro, text="그림 표[H]", command=그림용)
btn_pic.grid(row=3, column=2, columnspan=2, padx=5, pady=5)
bind_button_to_key(btn_pic, "h")

# '표 위치 매크로' 프레임 추가
frame1 = tk.Frame(
    tab1, highlightbackground="black", highlightthickness=2, padx=10, pady=10
)
frame1.grid(row=3, column=0, columnspan=2, padx=30, pady=7)

# 설명 레이블
position_description = tk.Label(
    frame1, text="표 위치 매크로", font=("Arial", 12, "bold")
)
position_description.grid(row=0, column=0, columnspan=4, padx=10, pady=7)

# 오른 단독 버튼
left_button = tk.Button(frame1, text="1단[V]", command=handle_1단용)
left_button.grid(row=1, column=3, rowspan=2, padx=10, pady=7)
bind_button_to_key(left_button, "v")

# 6개의 버튼 배열 (1열 3개, 2열 3개)
button1 = tk.Button(frame1, text="2단 왼 상[E]", command=handle_2단_왼_상)
button1.grid(row=1, column=0, padx=10, pady=7)
bind_button_to_key(button1, "e")

button2 = tk.Button(frame1, text="2단 가운 상[R]", command=handle_2단_가운_상)
button2.grid(row=1, column=1, padx=10, pady=7)
bind_button_to_key(button2, "r")

button3 = tk.Button(frame1, text="2단 오른 상[T]", command=handle_2단_오른_상)
button3.grid(row=1, column=2, padx=10, pady=7)
bind_button_to_key(button3, "t")

button4 = tk.Button(frame1, text="2단 왼 하[Z]", command=handle_2단_왼_하)
button4.grid(row=2, column=0, padx=10, pady=7)
bind_button_to_key(button4, "z")

button5 = tk.Button(frame1, text="2단 가운 하[X]", command=handle_2단_가운_하)
button5.grid(row=2, column=1, padx=10, pady=7)
bind_button_to_key(button5, "x")

button6 = tk.Button(frame1, text="2단 오른 하[C]", command=handle_2단_오른_하)
button6.grid(row=2, column=2, padx=10, pady=7)
bind_button_to_key(button6, "c")

# '표 위치 매크로' 프레임 추가
frame1 = tk.Frame(
    tab1, highlightbackground="black", highlightthickness=2, padx=10, pady=10
)
frame1.grid(row=0, column=2, columnspan=2, padx=30, pady=7)
# endregion

# region 나중에 만들거
########################################################
# 두 번째 탭
tab2 = ttk.Frame(notebook)
notebook.add(tab2, text="추가 기능 1")

frame2 = tk.Frame(tab2, padx=10, pady=10)
frame2.pack(expand=True, fill="both")
label_tab2 = tk.Label(
    frame2, text="여기에 두 번째 탭의 내용을 추가하세요.", font=("Arial", 12)
)
label_tab2.pack(pady=20)

# 세 번째 탭
tab3 = ttk.Frame(notebook)
notebook.add(tab3, text="추가 기능 2")

frame3 = tk.Frame(tab3, padx=10, pady=10)
frame3.pack(expand=True, fill="both")
label_tab3 = tk.Label(
    frame3, text="여기에 세 번째 탭의 내용을 추가하세요.", font=("Arial", 12)
)
label_tab3.pack(pady=20)
##############################################################
# endregion

# region UI 출력
root.update_idletasks()
root.geometry(
    f"{max(200, notebook.winfo_reqwidth())}x{max(300, notebook.winfo_reqheight() + 100)}"
)

# 파일 경로 업데이트
update_file_path_label()

# Tkinter 메인 루프 실행
root.mainloop()
# endregion
