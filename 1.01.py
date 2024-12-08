# %%
import tkinter as tk
from tkinter import ttk
from pyhwpx import Hwp

#region 초기화 부분
#인스턴스생성
hwp = Hwp()

# Tkinter 창 생성
root = tk.Tk()
root.title("송나겸 돈 복사 매크로(v1.01)")

#파일 경로 표시
file_path_frame = tk.Frame(root, highlightbackground="black", highlightthickness=1, padx=10, pady=10)
file_path_frame.grid(row=0, column=0, padx=15, pady=5, sticky="ew")

# 공지사항 라벨
notice_label = tk.Label(file_path_frame, text="2002방식 조판부호 사용 해야 작동함. / 창이 활성화된 상태에서만 단축키가 작동함.", font=("Arial", 10), anchor="w", fg="red")
notice_label.grid(row=1, column=0, columnspan=2, sticky="w", padx=5)  # 아래쪽 전체 너비로 표시

# 파일 경로 라벨
file_path_label = tk.Label(file_path_frame, text="파일 경로: 없음", font=("Arial", 10), anchor="e")
file_path_label.grid(row=0, column=0, sticky="w", padx=70)  # 왼쪽 정렬

#표 초기화
Table_list = [] # 모든 표의 갯수 찾기
Table_Total = 0
Table_index = 0 #전역으로 사용할 표 인덱스

# Notebook(탭 컨트롤러) 생성
notebook = ttk.Notebook(root)
notebook.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")  # pack 대신 grid 사용
tab1 = ttk.Frame(notebook)
notebook.add(tab1, text="표 셀/선 매크로")

# 탭의 열 비율 설정
tab1.grid_columnconfigure(0, weight=1, minsize=300)  # frame_table_movement가 있는 열
tab1.grid_columnconfigure(1, weight=3, minsize=300)  # frame_inspector가 있는 열

# '표 이동' 프레임
frame_table_movement = tk.Frame(tab1, highlightbackground="black", highlightthickness=2, padx=10, pady=10)
frame_table_movement.grid(row=0, column=0, pady=5, sticky="ew")

position_description = tk.Label(frame_table_movement, text="표 이동", font=("Arial", 12, "bold"))
position_description.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

align_ep_label = tk.Label(frame_table_movement, text=f"전체 표 갯수={Table_Total}")
align_ep_label.grid(row=0, column=2)

def Table_init():
    global Table_list
    global Table_Total
    global Table_index
    for i in hwp.ctrl_list:
        if i.UserDesc == "표":#컨트롤이 표일 경우
            Table_list.append(i)#리스트에 저장
    Table_Total = len(Table_list) # 문서안의 모든 표의 갯수
    print(Table_Total)
def update_file_path_label():
    """파일 경로를 업데이트하고 Label에 표시"""
    if hwp.Path:
        file_path = hwp.Path
        file_path_label.config(text=f"파일 경로: {file_path}")
        align_ep_label.config(text=f"전체 표 갯수={Table_Total}")
        
        print(f"파일 경로: {file_path}")
        print(f"전체 표 갯수={Table_Total}")
    else:
        file_path_label.config(text="파일 경로: 없음")
        print("파일 경로: 없음")

def select_file():
    """파일을 재선택"""
    try:
        #hwp.Save()
        #hwp.Run("FileClose")
        hwp.Run("FileOpen")
        Table_init()
        update_file_path_label()
        처음으로()
        root.update()  # 강제 업데이트
    except Exception as e:
        print(f"파일 재선택 중 오류: {e}")

# 파일 선택 버튼
file_reselect_button = tk.Button(file_path_frame, text="파일 선택", command=select_file)
file_reselect_button.grid(row=0, column=0, sticky="w", padx=5)  # 오른쪽 정렬
#endregion

#region 매크로 함수
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
    pset = hwp.HParameterSet.HCellBorderFill
    hwp.HAction.GetDefault("CellBorder", pset.HSet)
    pset.WidthVert = hwp.HwpLineWidth("0.12mm")
    pset.TypeVert = hwp.HwpLineType("Solid")
    pset.WidthHorz = hwp.HwpLineWidth("0.12mm")
    pset.TypeHorz = hwp.HwpLineType("Solid")
    hwp.HAction.Execute("CellBorder", pset.HSet)
def 표라인_위선(t : str):
    pset = hwp.HParameterSet.HCellBorderFill
    hwp.HAction.GetDefault("CellBorder", pset.HSet)
    pset.BorderWidthTop = hwp.HwpLineWidth(f"{t}mm")
    pset.BorderTypeTop = hwp.HwpLineType("Solid")
    return hwp.HAction.Execute("CellBorder", pset.HSet)
def 표라인_아래선(t : str):
    pset = hwp.HParameterSet.HCellBorderFill
    hwp.HAction.GetDefault("CellBorder", pset.HSet)
    pset.BorderWidthBottom = hwp.HwpLineWidth(f"{t}mm")
    pset.BorderTypeBottom = hwp.HwpLineType("Solid")
    return hwp.HAction.Execute("CellBorder", pset.HSet)
def 표라인_양옆투명():
    pset = hwp.HParameterSet.HCellBorderFill
    hwp.HAction.GetDefault("CellBorder", pset.HSet)
    pset.BorderWidthRight = hwp.HwpLineWidth("0.12mm")
    pset.BorderTypeRight = hwp.HwpLineType("None")
    pset.BorderWidthLeft = hwp.HwpLineWidth("0.12mm")
    pset.BorderTypeLeft = hwp.HwpLineType("None")
    pset.FillAttr.GradationAlpha = 0
    pset.FillAttr.ImageAlpha = 0
    hwp.HAction.Execute("CellBorder", pset.HSet)
                        
def 표색_없음():#안됨
    #hwp.cell_fill([255,255,255]) #흰색으로 채우기
    pset = hwp.HParameterSet.HCellBorderFill
    hwp.HAction.GetDefault("CellBorder", pset.HSet)
    pset.FillAttr.type = hwp.BrushType("NullBrush") # 소문자임.. ㅜㅜㅜㅜㅜ
    pset.FillAttr.GradationAlpha = 0
    pset.FillAttr.WindowsBrush = 0
    pset.FillAttr.ImageAlpha = 0
    return hwp.HAction.Execute("CellBorder", pset.HSet)#false,,,,,,,
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
    hwp.cell_fill([217,217,217]) 
def 표색_헤드_없음():
    hwp.TableColPageUp()
    pset = hwp.HParameterSet.HCellBorderFill
    hwp.HAction.GetDefault("CellBorder", pset.HSet)
    pset.FillAttr.type = hwp.BrushType("NullBrush") # 소문자임.. ㅜㅜㅜㅜㅜ
    pset.FillAttr.GradationAlpha = 0
    pset.FillAttr.WindowsBrush = 0
    pset.FillAttr.ImageAlpha = 0
    return hwp.HAction.Execute("CellBorder", pset.HSet)#false,,,,,,,
def 표라인_헤드_한줄():
    hwp.TableColPageUp()
    pset = hwp.HParameterSet.HCellBorderFill
    hwp.HAction.GetDefault("CellBorder", pset.HSet)
    pset.BorderWidthBottom = hwp.HwpLineWidth("0.12mm")
    pset.BorderTypeBottom = hwp.HwpLineType("Solid")
    hwp.HAction.Execute("CellBorder", pset.HSet)
def 표라인_헤드_두줄():
    hwp.TableColPageUp()
    pset = hwp.HParameterSet.HCellBorderFill
    hwp.HAction.GetDefault("CellBorder", pset.HSet)
    pset.BorderWidthBottom = hwp.HwpLineWidth("0.5mm")
    pset.BorderTypeBottom = hwp.HwpLineType("DoubleSlim")
    hwp.HAction.Execute("CellBorder", pset.HSet)
def 표라인_표주_윗선(t:str):
    pset = hwp.HParameterSet.HCellBorderFill
    hwp.HAction.Run("TableCellBlockRow")
    hwp.HAction.GetDefault("CellBorder", pset.HSet)
    pset.BorderWidthTop = hwp.HwpLineWidth(f"{t}mm")
    pset.BorderTypeTop = hwp.HwpLineType("Solid")
    pset.TypeVert = hwp.HwpLineType("None")
    pset.BorderTypeBottom = hwp.HwpLineType("None")
    pset.BorderTypeRight = hwp.HwpLineType("None")
    pset.BorderTypeLeft = hwp.HwpLineType("None")
    pset.FillAttr.GradationAlpha = 0
    pset.FillAttr.ImageAlpha = 0
    hwp.HAction.Execute("CellBorder", pset.HSet)
    hwp.TableCellBlockExtend()

def 표좌우맞춤():
    pagedef_dict = hwp.get_pagedef_as_dict()
    paper_width = int(pagedef_dict['용지폭'])       # 용지 폭
    margin_left = int(pagedef_dict['왼쪽'])         # 왼쪽 여백
    margin_right = int(pagedef_dict['오른쪽'])      # 오른쪽 여백
    # 실제 문서 폭 계산
    actual_width = paper_width - margin_left - margin_right
    epsilon = align_ep.get()
    hwp.set_table_width(actual_width-epsilon)
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
    #표 안쪽으로 들어가는 코드 추가
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
    #표 안쪽으로 들어가는 코드 추가
    hwp.MoveUp()
def 표위치_2단(Horz: str, Vert: str):
    """ Horz은  Center(왼),Left(가운데),Right(오른)  /
        Vert는 Top, Bottom
    """
    pset = hwp.HParameterSet.HShapeObject
    hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
    pset.HorzAlign = hwp.HAlign(Horz)
    pset.HorzRelTo = hwp.HorzRel("Page")
    pset.VertAlign = hwp.VAlign(Vert)
    pset.VertRelTo = hwp.VertRel("Page")
    pset.HorzOffset = hwp.MiliToHwpUnit(0.0)
    pset.VertOffset = hwp.MiliToHwpUnit(0.0)
    if Vert =="Top":
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
    """ Horz은  Center(왼),Left(가운데),Right(오른)  /
        Vert는 Top, Bottom
    """
    pset = hwp.HParameterSet.HShapeObject
    hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
    pset.HorzAlign = hwp.HAlign("Left")
    pset.HorzRelTo = hwp.HorzRel("Page")
    pset.VertAlign = hwp.VAlign("Top")
    pset.VertRelTo = hwp.VertRel("Para")
    pset.HorzOffset = hwp.MiliToHwpUnit(0.0)
    pset.VertOffset = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginBottom = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginTop = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginRight = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginLeft = hwp.MiliToHwpUnit(0.0)
    pset.HSet.SetItem("ShapeType", 3)
    pset.HSet.SetItem("ShapeCellSize", 0)
    hwp.HAction.Execute("TablePropertyDialog", pset.HSet)
#endregion

# region 함수 모음
def on_CellLineMacro(event=None):
    """버튼 클릭 이벤트: 드롭다운 값 출력"""
    윗선        = dropdown1.get()
    아래선      = dropdown2.get()
    양옆선      = dropdown3.get()
    내부선      = dropdown4.get()
    전체배경    = dropdown5.get()
    헤드배경    = dropdown6.get()
    헤드밑줄    = dropdown7.get()
    표주윗선    = dropdown8.get()

    if 윗선 == "냅둠":
        pass
    elif 윗선 == "0.12":
        표라인_위선("0.12")
    elif 윗선 == "0.4":
        표라인_위선("0.4")
    else:
        print(f"{윗선}이 선택되지 않았습니다.")    

    if 아래선 == "냅둠":
        pass
    elif 아래선 == "0.12":
        표라인_아래선("0.12")
    elif 아래선 == "0.4":
        표라인_아래선("0.4")
    else:
        print(f"{아래선}이 선택되지 않았습니다.")    

    if 양옆선 == "냅둠":
        pass
    elif 양옆선 == "투명":
        표라인_양옆투명()
    else:
        print(f"{양옆선}이 선택되지 않았습니다.")   
    
    if 내부선 == "냅둠":
        pass
    elif 내부선 == "0.12":
        표라인_안쪽실선()
    else:
        print(f"{내부선}이 선택되지 않았습니다.") 

    if 전체배경 == "냅둠":
        pass
    elif 전체배경 == "색없음":
        표색_없음()
    else:
        print(f"{전체배경}이 선택되지 않았습니다.")     

    if 헤드배경 == "냅둠":
        pass
    elif 헤드배경 == "없음":
        표색_헤드_없음()
        셀_전체선택()
    elif 헤드배경 == "회색(217)":
        표색_헤드_회색217()
        셀_전체선택()
    else:
        print(f"{헤드배경}이 선택되지 않았습니다.")  
    
    if 헤드밑줄 == "냅둠":
        pass
    elif 헤드밑줄 == "한줄":
        표라인_헤드_한줄()
    elif 헤드밑줄 == "두줄":
        표라인_헤드_두줄()
    else:
        print(f"{헤드밑줄}이 선택되지 않았습니다.")  

    if 표주윗선 == "없음(냅둠)":
        pass
    elif 표주윗선 == "0.12(아래투명/위는0.12)":
        표라인_표주_윗선("0.12")
    elif 표주윗선 == "0.4(아래투명/위0.4)":
        표라인_표주_윗선("0.4")
    else:
        print(f"{표주윗선}이 선택되지 않았습니다.")  

    
def 표여백정렬초기화():
    """안여백1,밖여백0,셀세로중앙정렬,안여백지정해제,글자처럼해제,자리차지,셀단위로나눔,제목줄반복,위캡션2미리"""
    hwp.set_table_inside_margin(1,1,1,1) #안여백 1mm로 밀기
    hwp.TableVAlignCenter() #셀 세로 중앙정렬
    안여백지정해제()
    글자처럼해제_자리차지()#1 위계가 있음.
    셀단위로나눔_제목줄반복()#2 위계가 있음.
    위캡션2mm()
    셀_전체선택()
    hwp.set_table_outside_margin(0,0,0,0)#밖여백 0mm로 밀기 
    셀_전체선택()
def 양옆맞추기(event=None):
    """양옆맞추기(오차0.4임)"""
    표좌우맞춤()
    #셀_전체선택ex()

def 단2맞추기(event=None):
    """2단일때 양옆 맞추기"""
    print("2단맞추기")

def 그림용(event=None):
    """그림 표용 세팅"""
    #표라인_전체투명()
    #표색_없음()###########################이거 따로 빼기
    #hwp.set_table_inside_margin(1,1,1,1) #안여벡 1mm로 밀기
    #hwp.TableVAlignCenter() #셀 세로 중앙정렬
    #안여백지정해제()
    #글자처럼해제_자리차지()#1 위계가 있음.
    #셀단위로나눔_제목줄반복()#2 위계가 있음.
    아래캡션3mm()
    셀_전체선택()
    #hwp.set_table_outside_margin(0,0,0,0)#밖여백 0mm로 밀기
    #셀_전체선택()

def handle_1단용(event=None):
    표위치_1단()
def handle_2단_왼_상(event=None):
    표위치_2단("Center","Top")
def handle_2단_가운_상(event=None):
    표위치_2단("Left","Top")
def handle_2단_오른_상(event=None):
    표위치_2단("Right","Top")

def handle_2단_왼_하(event=None):
    표위치_2단("Center","Bottom")
def handle_2단_가운_하(event=None):
    표위치_2단("Left","Bottom")
def handle_2단_오른_하(event=None):
    표위치_2단("Right","Bottom")

def 이전표(event=None):
    global Table_index
    global Table_list
    # Spinbox에서 현재 값을 가져오고 -1
    current_value = current_table_index.get()
    prev_index = max(0, current_value - 1)  # 최소값 제한
    current_table_index.set(prev_index)  # Spinbox에 값 설정
    Table_index = prev_index  # 전역 변수 업데이트
    print(f"현재 Table_index: {Table_index}")
    hwp.get_into_nth_table(Table_index)#인덱스 표의 첫번째 셀로 이동
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
    hwp.get_into_nth_table(Table_index)#인덱스 표의 첫번째 셀로 이동
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
            root.after(10, lambda: button.config(relief="raised"))  # 0.01초 후 원래 상태로 복구

    # 모든 키 입력을 감지하여 처리
    root.bind_all(f"<KeyPress-{key.lower()}>", key_action)
    root.bind_all(f"<KeyPress-{key.upper()}>", key_action)
#endregion

#region 첫번째 탭 : 표 셀/선 매크로

current_table_index = tk.IntVar(value=0)
align_ep = tk.DoubleVar(value=0.4)

##########################################
# 현재 표 인덱스와 이동 버튼
current_index_label = tk.Label(frame_table_movement, text="현재 표 인덱스:")
current_index_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")

index_spinbox = tk.Spinbox(frame_table_movement, from_=0, to=0, textvariable=current_table_index, width=5)
index_spinbox.grid(row=1, column=1, padx=5, pady=5)

btn_prev = tk.Button(frame_table_movement, text="이전 [Q]", command=이전표)
btn_prev.grid(row=2, column=0, padx=5, pady=5)
bind_button_to_key(btn_prev, "q")

btn_next = tk.Button(frame_table_movement, text="다음 [W]", command=다음표)
btn_next.grid(row=2, column=1, padx=5, pady=5)
bind_button_to_key(btn_next, "w")

btn_begin = tk.Button(frame_table_movement, text="처음으로 [B]", command=처음으로)
btn_begin.grid(row=1, column=2, columnspan=3, padx=5, pady=5)
bind_button_to_key(btn_begin, "b")
처음으로()

##########################################
# 'inspector' 프레임
frame_inspector = tk.Frame(tab1, highlightbackground="black", highlightthickness=2, padx=10, pady=10)
frame_inspector.grid(row=0, column=1, rowspan=4, padx=5, pady=5, sticky="nsew")

inspector_label = tk.Label(frame_inspector, text="표 inspector", font=("Arial", 12, "bold"))
inspector_label.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

##########################################
# '캡션/여백/정렬/글자해제' 프레임
frame_init_macro = tk.Frame(tab1, highlightbackground="black", highlightthickness=2, padx=10, pady=10)
frame_init_macro.grid(row=1, column=0,pady=5,sticky="ew")

init_macro_label = tk.Label(frame_init_macro, text="캡션/여백/정렬/글자해제", font=("Arial", 12, "bold")) # 기본 정렬 / 전체, 2단 채우기, 입실론, 그림 캡션
init_macro_label.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

btn_init = tk.Button(frame_init_macro, text="여백,표캡션,정렬 [E]", command=표여백정렬초기화)
btn_init.grid(row=1, column=0,  padx=5, pady=5,sticky="ew")
bind_button_to_key(btn_init, "e")

btn_pic = tk.Button(frame_init_macro, text="그림 캡션 [R]", command=그림용)
btn_pic.grid(row=1, column=1,padx=5, pady=5,sticky="ew")
bind_button_to_key(btn_pic, "r")

btn_align = tk.Button(frame_init_macro, text="전체 Fill [T]", command=양옆맞추기)
btn_align.grid(row=1, column=2, padx=5, pady=5,sticky="ew")
bind_button_to_key(btn_align, "t")

btn_align2 = tk.Button(frame_init_macro, text="2단 Fill(미완) [Y]", command=단2맞추기)
btn_align2.grid(row=1, column=3, padx=5, pady=5,sticky="ew")
bind_button_to_key(btn_align2, "y")

align_ep_label2 = tk.Label(frame_init_macro, text="Fill e =")
align_ep_label2.grid(row=0, column=2, padx=0, pady=5, sticky="e")

index_spinbox = tk.Spinbox(frame_init_macro, from_=-2, to=2, increment=0.1, textvariable=align_ep, width=3)
index_spinbox.grid(row=0, column=3, padx=0, pady=5, sticky="w")

##########################################
# '선/배경 설정' 프레임
frame_cell_macro = tk.Frame(tab1, highlightbackground="black", highlightthickness=2, padx=10, pady=10)
frame_cell_macro.grid(row=2, column=0,pady=5,sticky="ew")

cell_macro_label = tk.Label(frame_cell_macro, text="선/배경 설정", font=("Arial", 12, "bold"))#4개 받아서 바꾸기, 배경없음 선없음
cell_macro_label.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

#----------------------------------# 선택창

dropdown1_label = tk.Label(frame_cell_macro, text="윗 선:")
dropdown1_label.grid(row=1, column=0, padx=5, pady=5)

dropdown1 = ttk.Combobox(frame_cell_macro, values=["냅둠", "0.12", "0.4"])
dropdown1.set("0.4")
dropdown1.state(["readonly"])
dropdown1.grid(row=1, column=1, padx=5, pady=5)

dropdown2_label = tk.Label(frame_cell_macro, text="아래 선:")
dropdown2_label.grid(row=2, column=0, padx=5, pady=5)

dropdown2 = ttk.Combobox(frame_cell_macro, values=["냅둠", "0.12", "0.4"])
dropdown2.set("0.4")
dropdown2.state(["readonly"])
dropdown2.grid(row=2, column=1, padx=5, pady=5)

dropdown3_label = tk.Label(frame_cell_macro, text="양옆 선:")
dropdown3_label.grid(row=3, column=0, padx=5, pady=5)

dropdown3 = ttk.Combobox(frame_cell_macro, values=["냅둠","투명"])
dropdown3.set("투명")
dropdown3.state(["readonly"])
dropdown3.grid(row=3, column=1, padx=5, pady=5)

dropdown4_label = tk.Label(frame_cell_macro, text="내부 선:")
dropdown4_label.grid(row=4, column=0, padx=5, pady=5)

dropdown4 = ttk.Combobox(frame_cell_macro, values=["냅둠","0.12"])
dropdown4.set("냅둠")
dropdown4.state(["readonly"])
dropdown4.grid(row=4, column=1, padx=5, pady=5)

dropdown5_label = tk.Label(frame_cell_macro, text="전체 배경:")
dropdown5_label.grid(row=5, column=0, padx=5, pady=5)

dropdown5 = ttk.Combobox(frame_cell_macro, values=["냅둠","색없음"])
dropdown5.set("냅둠")
dropdown5.state(["readonly"])
dropdown5.grid(row=5, column=1, padx=5, pady=5)

dropdown6_label = tk.Label(frame_cell_macro, text="헤드 배경:")
dropdown6_label.grid(row=6, column=0, padx=5, pady=5)

dropdown6 = ttk.Combobox(frame_cell_macro, values=["냅둠", "없음", "회색(217)"])
dropdown6.set("냅둠")
dropdown6.state(["readonly"])
dropdown6.grid(row=6, column=1, padx=5, pady=5)

dropdown7_label = tk.Label(frame_cell_macro, text="헤드 밑 줄:")
dropdown7_label.grid(row=7, column=0, padx=5, pady=5)

dropdown7 = ttk.Combobox(frame_cell_macro, values=["냅둠","한줄", "두줄"])
dropdown7.set("냅둠")
dropdown7.state(["readonly"])
dropdown7.grid(row=7, column=1, padx=5, pady=5)

dropdown8_label = tk.Label(frame_cell_macro, text="표주 윗선:")
dropdown8_label.grid(row=7, column=0, padx=5, pady=5)

dropdown8 = ttk.Combobox(frame_cell_macro, values=["없음(냅둠)","0.12(아래투명/위는0.12)", "0.4(아래투명/위0.4)"])
dropdown8.set("없음(냅둠)")
dropdown8.state(["readonly"])
dropdown8.grid(row=7, column=1, padx=5, pady=5)

#-------------------#


btn_execute = tk.Button(frame_cell_macro, text="<<< 적용 [F]", command=on_CellLineMacro)
btn_execute.grid(row=1, column=2, padx=5, pady=5,sticky="ew")
bind_button_to_key(btn_execute, "f")

btn_execute = tk.Button(frame_cell_macro, text="All배경없음 [G]", command=표색_없음)
btn_execute.grid(row=6, column=2,  padx=5, pady=5,sticky="e")
bind_button_to_key(btn_execute, "g")

btn_execute = tk.Button(frame_cell_macro, text="  All선없음 [h]", command=표라인_전체투명)
btn_execute.grid(row=7, column=2,  padx=5, pady=5,sticky="e")
bind_button_to_key(btn_execute, "h")

###################

##########################################
# '표 위치 설정' 프레임 추가
frame1 = tk.Frame(tab1, highlightbackground="black", highlightthickness=2, padx=10, pady=10)
frame1.grid(row=3, column=0, pady=5, sticky="ew")

# 설명 레이블
position_description = tk.Label(frame1, text="표 위치 매크로", font=("Arial", 12, "bold"))
position_description.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

# 오른 단독 버튼
left_button = tk.Button(frame1, text="1단[V]", command=handle_1단용)
left_button.grid(row=1, column=3, rowspan=2, padx=10, pady=7)
bind_button_to_key(left_button, "v")

# 6개의 버튼 배열 (1열 3개, 2열 3개)
button1 = tk.Button(frame1, text="2단 왼 상[A]", command=handle_2단_왼_상)
button1.grid(row=1, column=0, padx=10, pady=7)
bind_button_to_key(button1, "a")

button2 = tk.Button(frame1, text="2단 가운 상[S]", command=handle_2단_가운_상)
button2.grid(row=1, column=1, padx=10, pady=7)
bind_button_to_key(button2, "s")

button3 = tk.Button(frame1, text="2단 오른 상[D]", command=handle_2단_오른_상)
button3.grid(row=1, column=2, padx=10, pady=7)
bind_button_to_key(button3, "d")

button4 = tk.Button(frame1, text="2단 왼 하[Z]", command=handle_2단_왼_하)
button4.grid(row=2, column=0, padx=10, pady=7)
bind_button_to_key(button4, "z")

button5 = tk.Button(frame1, text="2단 가운 하[X]", command=handle_2단_가운_하)
button5.grid(row=2, column=1, padx=10, pady=7)
bind_button_to_key(button5, "x")

button6 = tk.Button(frame1, text="2단 오른 하[C]", command=handle_2단_오른_하)
button6.grid(row=2, column=2, padx=10, pady=7)
bind_button_to_key(button6, "c")
##########################################

#endregion

#region 나중에 만들거
########################################################
# 두 번째 탭
tab2 = ttk.Frame(notebook)
notebook.add(tab2, text="추가 기능 1")

frame2 = tk.Frame(tab2, padx=10, pady=10)
frame2.pack(expand=True, fill="both")
label_tab2 = tk.Label(frame2, text="여기에 두 번째 탭의 내용을 추가하세요.", font=("Arial", 12))
label_tab2.pack(pady=20)

# 세 번째 탭
tab3 = ttk.Frame(notebook)
notebook.add(tab3, text="추가 기능 2")

frame3 = tk.Frame(tab3, padx=10, pady=10)
frame3.pack(expand=True, fill="both")
label_tab3 = tk.Label(frame3, text="여기에 세 번째 탭의 내용을 추가하세요.", font=("Arial", 12))
label_tab3.pack(pady=20)
##############################################################
#endregion

#region UI 출력
root.update_idletasks()
root.geometry(f"{max(200, 15+ notebook.winfo_reqwidth())}x{max(300, notebook.winfo_reqheight() + 100)}")

# 파일 경로 업데이트
update_file_path_label()

# Tkinter 메인 루프 실행
root.mainloop()
#endregion



