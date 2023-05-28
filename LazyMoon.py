import tkinter as tk

class PopUp(tk.Tk):
    def __init__(self, callback):
        super().__init__()
        self.title("이미지 크기")
        self.geometry("250x100") 
        self.resizable(False, False)  
        self.iconbitmap('./.ico/me.ico')
        self.callback = callback
        print("Asd")

        self.Xentry = tk.Entry(self, width=15)
        self.Yentry = tk.Entry(self, width=15)

        tk.Label(self, text="너비(mm) : ", width=15).grid(row=0, column=0, pady=5)
        tk.Label(self, text="높이(mm) : ", width=15).grid(row=1, column=0)
        button = tk.Button (self, text='확인',command=self.get_WH, width=20)
        
        self.Xentry.grid(row=0, column=1, pady=5)
        self.Yentry.grid(row=1, column=1)
        button.grid(row=2, column=0, columnspan=2, pady=10)

    def get_WH(self):
        try:
            xW = float(self.Xentry.get())
            yH = float(self.Yentry.get())
            if xW > 0 and yH > 0:
                self.callback(xW, yH)
                self.destroy()
            else:
                raise(ValueError)
        except(ValueError):
            import tkinter.messagebox
            tkinter.messagebox.showerror(    "메시지 알림", "  0보다 큰 수를 입력하세요!!\n\n    다시 시도하세요")
            self.destroy()


class LazyMoon(tk.Tk):

    image_dict = dict()
    file_name = str()
    xW = int()
    yH = int()

    def __init__(self):
        super().__init__()

        self.title(" LazyMoon")
        self.geometry("400x300+50+50")
        self.resizable(False, False)
        self.iconbitmap('./.ico/me.ico')

        text = tk.Label(self, text="이 프로그램은 한글과 엑셀 파일에 사진을 \n자동으로 첨부해주는 자동문서화 프로그램입니다.\n")
        btn1 = tk.Button(self, text="hwp, xlsx 파일 선택", command=self.get_file, width=30)
        btn2 = tk.Button(self, text="이미지 선택", command=self.get_imgList, width=30)
        btn3 = tk.Button(self, text="이미지 크기 설정", command=self.setPopUp, width=30)
        btn4 = tk.Button(self, text="실행하기", command=self.start, width=30, height=2)
        text2 = tk.Label(self, text="만든 사람 : https://github.com/mb5ss95", height=2)

        text.pack(pady="10")
        btn1.pack(side="top", pady="5")
        btn2.pack(side="top", pady="5")
        btn3.pack(side="top", pady="5")
        btn4.pack(side="top", pady="5")
        text2.pack(side="top", pady="10")


    def exel_find_insert(self, ws, direction, name, msg):
        import os
        from openpyxl.drawing.image import Image

        for row, values in enumerate(ws.values):
            for col, value in enumerate(values):
                if value == msg:
                    p = ws.cell(row=row+1, column=col+1, value="")
                    w = p.column_letter+str(p.row)
                    ws.row_dimensions[p.row].height = round(self.yH*2.834)
                    ws.column_dimensions[p.column_letter].width = round(self.xW*0.457)
                    img = Image(os.path.join(direction, name))
                    print(img.height, img.width)
                    img.height = round(self.yH/0.26458)
                    img.width = round(self.xW/0.26458)
                    #220ppi
                    ws.add_image(img, w)
                
                    print("EXCEL IMG : ", name, ", ", msg, ", ", self.xW, ", ", self.yH, ", ", w, ", ", value)



    def hwp_find_insert(self, hwp, direction, name, msg):
        import os

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
                print("HWP IMG : ", name, ", ", msg, ", ", self.xW, ", ", self.yH)
                #INSERT IMG :  스크린샷(1).png ,  $1$ ,  40 ,  40
                
                hwp.InsertPicture(os.path.join(direction, name), Embedded=True, Width=self.xW, Height=self.yH, sizeoption=1)
            else:
                break


    def start(self):
        from tkinter import messagebox as msg

        if not bool(self.file_name):
            msg.showwarning('메시지 알림', '파일을 넣어주세요!!')
            if not bool(self.image_dict):  
                msg.showwarning('메시지 알림', '이미지 파일을 넣어주세요!!') 
            if not bool(self.xW) or not bool(self.yH):
                msg.showwarning('메시지 알림', '이미지 크기를 설정하세요!!') 
                return
            return
        else:
            if not bool(self.image_dict):  
                msg.showwarning('메시지 알림', '이미지 파일을 넣어주세요!!')
                if not bool(self.xW) or not bool(self.yH):
                    msg.showwarning('메시지 알림', '이미지 크기를 설정하세요!!') 
                    return 
                return

        print(self.image_dict)
        #{'direction': 'C:/Users/moon/Pictures/Screenshots', 'name': ['스크린샷(1).png', '스크린샷(2).png']}
        
        img_direction = self.image_dict['direction']
        img_list = self.image_dict['name']

        if self.file_name[-4:] == ".hwp":
            import win32com.client as win32
            
            hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
            hwp.Open(self.file_name, "HWP", "forceopen:true")

            for index, img_name in enumerate(img_list):
                self.hwp_find_insert(hwp, img_direction, img_name, img_name)
                self.hwp_find_insert(hwp, img_direction, img_name, "$"+str(index+1)+"$")
        else:
            import openpyxl
            import tkinter.messagebox

            try:
                wb = openpyxl.load_workbook(self.file_name)
                ws = wb.active
                #print(wb.get_sheet_names())
                
                for index, img_name in enumerate(img_list):
                    self.exel_find_insert(ws, img_direction, img_name, img_name)
                    self.exel_find_insert(ws, img_direction, img_name, "$"+str(index+1)+"$")
                
                wb.save(self.file_name)
                tkinter.messagebox.showinfo(    "메시지 알림", "%s로 저장 되었습니다!"%self.file_name)
            except(PermissionError):
                tkinter.messagebox.showerror(    "메시지 알림", "  엑셀 파일을 닫고 실행해주세요!!\n\n    다시 시도하세요")

        self.destroy()
        # hwp.Clear(3)
        # hwp.Quit()


    def get_xy(self, *args):
        self.xW = args[0]
        self.yH = args[1]

    def setPopUp(self):
        print("Asd")
        popUp = PopUp(self.get_xy)
        popUp.mainloop()


    def get_imgList(self):
        from tkinter.filedialog import askopenfilenames

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
            self.image_dict = {'direction': direc_list, 'name': image_list}

            from tkinter import messagebox
            messagebox.showinfo(title="이미지 목록", message=str(image_list)+"\n")

        except IndexError:
            return


    def get_file(self):
        from tkinter.filedialog import askopenfilenames

        hwpFile = askopenfilenames(title='이미지를 삽입할 파일을 선택하세요', filetypes=[
            ('모든 문서 파일', '*.hwp *.xlsx *.xlsm'),
            ('한글 파일 (.hwp)', '*.hwp'), 
            ('엑셀 파일 (.xlsx .xlsm)', '*.xlsx *.xlsm')])

        try:
            self.file_name = hwpFile[0]
        except IndexError:
            return


if __name__ == '__main__':
    lazyMoon = LazyMoon()
    lazyMoon.mainloop()