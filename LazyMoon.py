from tkinter import XView


image_dict = dict()
file_name = str()
xW = int()
yH = int()


def exel_find_insert(ws, direction, name, msg):
    import os
    from openpyxl.drawing.image import Image

    global xW,yH

    #ws.add_image(img, 'A2')
    for row, values in enumerate(ws.values):
        for col, value in enumerate(values):
            if value == msg:
                p = ws.cell(row=row+1, column=col+1, value="")
                w = p.column_letter+str(p.row)
                ws.row_dimensions[p.row].height = round(yH*2.834)
                ws.column_dimensions[p.column_letter].width = round(xW*0.457)
                img = Image(os.path.join(direction, name))
                print(img.height, img.width)
                img.height = round(yH/0.26458)
                img.width = round(xW/0.26458)
                #220ppi
                ws.add_image(img, w)
            
                print("EXCEL IMG : ", name, ", ", msg, ", ", xW, ", ", yH, ", ", w, ", ", value)



def hwp_find_insert(hwp, direction, name, msg):
    import os

    global xW,yH

    hwp.MovePos(2, 0, 0)
    hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)

    option = hwp.HParameterSet.HFindReplace
    option.FindString = msg
    option.UseWildCards = 1
    option.IgnoreMessage = 1
    option.Direction = hwp.FindDir("Forward")
    option.FindType = False

    while(1):
        if hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet):
            print("HWP IMG : ", name, ", ", msg, ", ", xW, ", ", yH)
            #INSERT IMG :  스크린샷(1).png ,  $1$ ,  40 ,  40
            
            hwp.InsertPicture(os.path.join(direction, name), Embedded=True, Width=xW, Height=yH, sizeoption=1)
        else:
            break


def start():
    from tkinter import messagebox as msg

    global image_dict, window, file_name, xW, yH

    if not bool(file_name):
        msg.showwarning('메시지 알림', '파일을 넣어주세요!!')
        if not bool(image_dict):  
            msg.showwarning('메시지 알림', '이미지 파일을 넣어주세요!!') 
        if not bool(xW) or not bool(yH):
            msg.showwarning('메시지 알림', '이미지 크기를 설정하세요!!') 
            return
        return
    else:
        if not bool(image_dict):  
            msg.showwarning('메시지 알림', '이미지 파일을 넣어주세요!!')
            if not bool(xW) or not bool(yH):
                msg.showwarning('메시지 알림', '이미지 크기를 설정하세요!!') 
                return 
            return
    #  ('hwp files', '*.hwp'), ('exel files', '*.xlsx *.xlsm')

    print(image_dict)
    #{'direction': 'C:/Users/moon/Pictures/Screenshots', 'name': ['스크린샷(1).png', '스크린샷(2).png']}
    
    img_direction = image_dict['direction']
    img_list = image_dict['name']

    if file_name[-4:] == ".hwp":
        import win32com.client as win32
        
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.Open(file_name, "HWP", "forceopen:true")

        for index, img_name in enumerate(img_list):
            hwp_find_insert(hwp, img_direction, img_name, img_name)
            hwp_find_insert(hwp, img_direction, img_name, "$"+str(index+1)+"$")
    else:
        import openpyxl
        import tkinter.messagebox

        try:
            wb = openpyxl.load_workbook(file_name)
            ws = wb.active
            #print(wb.get_sheet_names())
            
            for index, img_name in enumerate(img_list):
                exel_find_insert(ws, img_direction, img_name, img_name)
                exel_find_insert(ws, img_direction, img_name, "$"+str(index+1)+"$")
            
            wb.save(file_name)
            tkinter.messagebox.showinfo(    "메시지 알림", "%s로 저장 되었습니다!"%file_name)
        except(PermissionError):
            tkinter.messagebox.showerror(    "메시지 알림", "  엑셀 파일을 닫고 실행해주세요!!\n\n    다시 시도하세요")

    window.destroy()
    # hwp.Clear(3)
    # hwp.Quit()


def get_xy():
    import tkinter as tk
    from tkinter import messagebox

    root = tk.Tk()
    root.title("이미지 크기")
    root.geometry("250x100") 
    root.resizable(False, False)  
    root.iconbitmap('./.ico/DSU.ico')

    tk.Label(root, text="너비(mm) : ", width=15).grid(row=0, column=0, pady=5)
    Xentry = tk.Entry(root, width=15)
    Xentry.grid(row=0, column=1, pady=5)

    tk.Label(root, text="높이(mm) : ", width=15).grid(row=1, column=0)
    Yentry = tk.Entry(root, width=15)
    Yentry.grid(row=1, column=1)

    def get_WH():
        global xW, yH

        try:
            xW = float(Xentry.get())
            yH = float(Yentry.get())
            if xW > 0 and yH > 0:
                root.destroy()
            else:
                raise(ValueError)
        except(ValueError):
            import tkinter.messagebox
            tkinter.messagebox.showerror(    "메시지 알림", "  0보다 큰 수를 입력하세요!!\n\n    다시 시도하세요")
            root.destroy()

    button = tk.Button (root, text='확인',command=get_WH, width=20)
    button.grid(row=2, column=0, columnspan=2, pady=10)
    
    root.mainloop()


def get_imgList():
    from tkinter.filedialog import askopenfilenames

    global image_dict

    imagelist = askopenfilenames(
        title='이미지를 선택하세요', filetypes=[
            ('모든 그림 파일', '*.bmp *.cdr *.drw *.dxf *.emf *.gif *.jpg *.jpeg *.pcx *.pic *.png *.svg *.tif *.wmf'),
            ('BMP (.bmp)', '*.bmp'), 
            ('CDR (.cdr)', '*.cdr'), 
            ('DRW (.drw)', '*.drw'), 
            ('DXF (.dxf)', '*.dxf'), 
            ('EMF (.emf)', '*.emf'), 
            ('GIF (.gif)', '*.gif'),
            ('JPG (.jpg .jpeg)', '*.jpg *.jpeg'),
            ('PCX (.pcx)', '*.pcx'),
            ('PIC (.pic)', '*.pic'),
            ('PNG (.png)', '*.png'),
            ('SVG (.svg)', '*.svg'),
            ('TIFF (.tif)', '*.tif'),
            ('WMF (.wmf)', '*.wmf')])

    try:
        direc_list = imagelist[0].rsplit("/", maxsplit=1)[0]
        image_list = [i.rsplit("/", maxsplit=1)[1] for i in imagelist]
        image_dict = {'direction': direc_list, 'name': image_list}

        from tkinter import messagebox
        messagebox.showinfo(title="이미지 목록", message=str(image_list)+"\n")

    except IndexError:
        return


def get_file():
    from tkinter.filedialog import askopenfilenames

    global file_name

    hwpFile = askopenfilenames(title='이미지를 삽입할 파일을 선택하세요', filetypes=[
        ('모든 문서 파일', '*.hwp *.xlsx *.xlsm'),
        ('한글 파일 (.hwp)', '*.hwp'), 
        ('엑셀 파일 (.xlsx .xlsm)', '*.xlsx *.xlsm')])

    try:
        file_name = hwpFile[0]
    except IndexError:
        return


if __name__ == '__main__':
    import tkinter as tk
    
    window = tk.Tk()
    window.title(" LazyMoon 2.0")
    window.geometry("400x300+50+50")
    window.resizable(False, False)
    window.iconbitmap('./.ico/DSU.ico')

    text = tk.Label(window, text="이 프로그램은 한글과 엑셀 파일에 사진을 \n자동으로 첨부해주는 자동문서화 프로그램입니다.\n")
    text.pack(pady="10")

    btn1 = tk.Button(window, text="hwp, xlsx 파일 선택", command=get_file, width=30)
    btn1.pack(side="top", pady="5")

    btn2 = tk.Button(window, text="이미지 선택", command=get_imgList, width=30)
    btn2.pack(side="top", pady="5")

    btn3 = tk.Button(window, text="이미지 크기 설정", command=get_xy, width=30)
    btn3.pack(side="top", pady="5")

    btn4 = tk.Button(window, text="실행하기", command=start, width=30, height=2)
    btn4.pack(side="top", pady="5")

    text2 = tk.Label(window, text="만든 사람 : 동서울대학교 MBS", height=2)
    text2.pack(side="top", pady="10")

    window.mainloop()
