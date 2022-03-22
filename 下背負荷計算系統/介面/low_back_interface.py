# -*- coding: utf-8 -*-
"""
Created on Wed Mar  4 17:51:58 2020

@author: 高漢佑
"""
import Inverse_dynamic_model as idm
import kinect as k45
from tkinter import messagebox, Canvas
from tkinter import ttk
import tkinter as tk, threading   
from tkinter import font  as tkfont 
import imageio
from PIL import Image, ImageTk
import time
from datetime import date
from tkcalendar import Calendar
from tkinter import filedialog as fd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.patches import FancyArrowPatch
from matplotlib import animation
from matplotlib.figure import Figure
from mpl_toolkits.mplot3d import proj3d
import statistics 
import mpl_toolkits.mplot3d as plt3d
import numpy as np

l5s1_data = []
kinect1 = []
joint_data = []
joint_plot = [[1,13],[1,34],[28,34],[4,16],[4,37],[31,37],[28,46],[31,46],[46,49],[25,49],[19,52],[7,52],[7,40],[40,49],[22,55],[10,55],[10,43],[43,49]]
joint_plot2 = [[0,6],[6,12],[12,18],[18,24],[3,9],[9,15],[15,21],[21,24],[24,27],[27,30],[27,33],[27,36],[33,39],[36,42],[39,45],[42,48],[45,51],[48,54]]
class SampleApp(tk.Tk):
    
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        self.title_font = tkfont.Font(family='Helvetica', size=18, weight="bold", slant="italic")  
        container = tk.Frame(self,background='white')
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        self.shared_data = {"username": None,"year":None,
                            "height":None,"weight":None,
                            "sex":None,"origin":None,
                            "end":None,"box":None,"date": date.today(),
                            "method":"0"}
        self.frames = {}
        for F in (StartPage, PageOne, PageTwo, PageThree, PageFour1, PageFour2, PageFour3, PageFour4, PageProcess, PageFive1, PageFive2):
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame            
            frame.grid(row=0, column=0, sticky="nsew")
        self.show_frame("StartPage")

    def show_frame(self, page_name):            
        frame = self.frames[page_name]
        frame.tkraise()
        
    

class StartPage(tk.Frame):

    def __init__(self, parent, controller):
        def sys_quit():            
            app.destroy()
            label1.config(text="")            
        tk.Frame.__init__(self, parent)
        self.controller = controller
        self.configure(bg='white')
        def pause():                                                           #註解80-87
            time.sleep(0.05)
            pass
        def stream(label):                         
            for images in video.iter_data():     
                pause()                
                frame_image = ImageTk.PhotoImage(image = Image.fromarray(images))                
                label.config(image=frame_image)
                label.image = frame_image
        label1 = tk.Label(self,text="人工抬舉下背生物力學分析系統",font = "微軟正黑體 20 bold").pack(pady=10)
        video_name = "mp4/icon.mp4" 
        video = imageio.get_reader(video_name)
        label = tk.Label(self,width=240, height=240)
        label.config(bg="white")
        label.pack(pady=20)
        thread = threading.Thread(target=stream, args=(label,))
        thread.daemon = 1
        thread.start()    
        button1 = tk.Button(self,text="關閉系統",font = "微軟正黑體 14",command=sys_quit,bg='gray',fg='snow')        
        button2 = tk.Button(self, text="進入系統",font = "微軟正黑體 14",command=lambda: controller.show_frame("PageOne"),bg='dodger blue',fg='snow')        
        button1.pack(padx=100,pady=5,side='left',ipadx=2)
        button2.pack(padx=100,pady=5,side='right',ipadx=2)

class PageOne(tk.Frame):

    def __init__(self, parent, controller):
        def get_data():            
            self.controller.shared_data["username"] = nameString.get()
            self.controller.shared_data["year"] = yearString.get()
            self.controller.shared_data["height"] = heightString.get()
            self.controller.shared_data["weight"] = weightString.get()
            self.controller.shared_data["sex"] = self.v.get()  
            
            controller.show_frame("PageTwo")
        tk.Frame.__init__(self, parent)
        self.controller = controller
        frame = tk.LabelFrame(self, text = "基本資料", padx=2,pady=2,labelanchor = "n",font='微軟正黑體 14 bold')
        frame.place(x=330,y=40)
        label2 = tk.Label(frame, text = "工作人員",font=('微軟正黑體',10))
        label2.pack(pady=2)
        nameString = tk.StringVar()
        self.entry = tk.Entry(frame, textvariable = nameString)
        self.entry.pack(padx=2,pady=2)
        label4 = tk.Label(frame, text = "年齡",font=('微軟正黑體',10)).pack(pady=5)
        yearString = tk.StringVar()
        entry2 = tk.Entry(frame, textvariable = yearString)
        entry2.pack(padx=2,pady=2)
        heightString = tk.StringVar()
        label3 = tk.Label(frame, text = "身高 (cm)",font=('微軟正黑體',10))
        label3.pack(pady=5)
        entry3 = tk.Entry(frame, textvariable = heightString)
        entry3.pack(padx=2,pady=2)
        weightString = tk.StringVar()
        label4 = tk.Label(frame, text = "體重 (kg)",font=('微軟正黑體',10))
        label4.pack(pady=5)
        entry4 = tk.Entry(frame, textvariable = weightString)
        entry4.pack(padx=2,pady=2)
        label5 = tk.Label(frame, text = "性別",font=('微軟正黑體',10))
        label5.pack(pady=5)
        self.v = tk.StringVar()
        self.v.set(1)
        sex_item1 = tk.Radiobutton(frame, text="男性", value="1", variable=self.v,font=('微軟正黑體',10))
        sex_item1.pack(padx=2,fill='x',side='left')
        sex_item2 = tk.Radiobutton(frame, text="女性", value="0", variable=self.v,font=('微軟正黑體',10))
        sex_item2.pack(padx=2,fill='x',side='right')
        photo = tk.PhotoImage(file = "img/right.png")
        photoimage = photo.subsample(6, 6) 
        photo2 = tk.PhotoImage(file = "img/left.png")
        photoimage2 = photo2.subsample(6, 6) 
        button1 = tk.Button(self,image = photoimage,command=lambda: get_data(),padx=5,pady=5)
        button1.image = photoimage
        button2 = tk.Button(self, image = photoimage2,command=lambda: controller.show_frame("StartPage"),padx=5,pady=5)
        button2.image = photoimage2
        button1.place(x=757,y=150)
        button2.place(x=0,y=150)        

class PageTwo(tk.Frame):

    def __init__(self, parent, controller):
        def get_data():            
            self.controller.shared_data["origin"] = c1.get()
            self.controller.shared_data["end"] = c2.get()
            self.controller.shared_data["date"] = text.get()
            self.controller.shared_data["box"] = boxString.get()
            controller.show_frame("PageThree")
        def cal_func():
            def select_date():
                text.set(cal.get_date()) 
                top.destroy()
            today = date.today()       
            top = tk.Toplevel(self)
            cal = Calendar(top, font = "微軟正黑體 14", selectmode = "day", year = today.year, month = today.month, day = today.day, weekendforeground ='red', showweeknumbers = False, date_pattern='y-mm-dd')
            cal.pack(fill = "both", expand = True)
            btn3 = tk.Button(top, text = "OK",command = select_date)
            btn3.pack()
        
        tk.Frame.__init__(self, parent)
        self.controller = controller
        skeleton_img = tk.PhotoImage(file = "img/skeleton.png")
        skeleton_img = skeleton_img.subsample(2, 2) 
        skeleton_label = tk.Label(self, padx=5,pady=10, image = skeleton_img,height=314,width=300)
        skeleton_label.image = skeleton_img
        skeleton_label.place(x=130,y=40)
        frame2 = tk.LabelFrame(self, text = "動作資訊", padx=5,pady=10, labelanchor="n",font='微軟正黑體 14 bold',height=314,width=300)
        frame2.place(x=470,y=50)
        label = tk.Label(frame2, text = "搬運起點",font=('微軟正黑體',10))
        label.pack(pady=2)
        c1=tk.StringVar()
        comboxlist=ttk.Combobox(frame2,textvariable=c1) #初始化
        comboxlist["values"]=("地板","手肘","膝蓋","指節","腰部","肩膀","頭部")
        comboxlist.pack(pady=2)
        label2 = tk.Label(frame2, text = "搬運終點",font=('微軟正黑體',10))
        label2.pack(pady=2)
        c2=tk.StringVar()
        comboxlist2=ttk.Combobox(frame2,textvariable=c2) #初始化
        comboxlist2["values"]=("地板","手肘","膝蓋","指節","腰部","肩膀","頭部")
        comboxlist2.pack(pady=2)  
        label3 = tk.Label(frame2, text = "箱子重量 (kg)",font=('微軟正黑體',10))
        label3.pack(padx=2,pady=1)
        boxString = tk.StringVar()
        entry = tk.Entry(frame2, textvariable = boxString)
        entry.pack(padx=2,pady=2)
        label4 = tk.Label(frame2, text = "搬運日期",font=('微軟正黑體',10))
        label4.pack(padx=2,pady=1)
        cal = tk.PhotoImage(file = "img/calendar.png")
        text = tk.StringVar()
        text.set(date.today())
        photoimage = cal.subsample(20, 20) 
        btn1 = tk.Button(frame2, text = "日期",command = cal_func, image = photoimage)
        btn1.image = photoimage
        btn1.pack()
        label5 = tk.Label(frame2, textvariable = text,font=('微軟正黑體',10))
        label5.pack(padx=5,pady=1)
        photo = tk.PhotoImage(file = "img/right.png")
        photoimage = photo.subsample(6, 6) 
        photo2 = tk.PhotoImage(file = "img/left.png")
        photoimage2 = photo2.subsample(6, 6) 
        button1 = tk.Button(self,image = photoimage,command=lambda: get_data(),padx=5,pady=5)
        button1.image = photoimage
        button2 = tk.Button(self, image = photoimage2,command=lambda: controller.show_frame("PageOne"),padx=5,pady=5)
        button2.image = photoimage2
        button1.place(x=757,y=150)
        button2.place(x=0,y=150) 


class PageThree(tk.Frame):
    
    def __init__(self, parent, controller):   
        tk.Frame.__init__(self, parent)
        self.controller = controller        
        label = tk.Label(self, text="評估工具",font='微軟正黑體 14 bold')        
        label.pack(pady=2)
        self.method = tk.StringVar()
        self.method.set(0)         
        def get_data():
            self.controller.shared_data["method"] = self.method.get()            
            if self.controller.shared_data["method"] == '0': 
                self.controller.frames["PageFour1"].correct_label()
                controller.show_frame("PageFour1")  
            elif self.controller.shared_data["method"] == '1':
                self.controller.frames["PageFour2"].correct_label()
                controller.show_frame("PageFour2")
            elif self.controller.shared_data["method"] == '2':
                self.controller.frames["PageFour3"].correct_label()
                controller.show_frame("PageFour3")
            else:
                self.controller.frames["PageFour4"].correct_label()
                controller.show_frame("PageFour4")
            
        motion_img = tk.PhotoImage(file = "img/mts.png")
        motion_img = motion_img.subsample(4, 4) 
        labe2 = tk.Label(self, image = motion_img,height=180,width=195)
        labe2.image = motion_img
        labe2.place(x=80,y=80)
        kinect1_img = tk.PhotoImage(file = "img/k.png")
        kinect1_img = kinect1_img.subsample(3, 3) 
        label3 = tk.Label(self, image = kinect1_img,height=180,width=195)
        label3.image = kinect1_img
        label3.place(x=300,y=80)
        #kinect2_img = tk.PhotoImage(file = "img/kinect2.png")
        #kinect2_img = kinect2_img.subsample(2, 2) 
        label4 = tk.Label(self, image = kinect1_img,height=180,width=195)
        label4.image = kinect1_img
        label4.place(x=520,y=80)
        #kinect3_img = tk.PhotoImage(file = "img/kinect3.png")
        #kinect3_img = kinect3_img.subsample(2, 2) 
        #label5 = tk.Label(self, image = kinect3_img,height=180,width=195)
        #label5.image = kinect3_img
        #label5.place(x=420,y=190)
        method1 = tk.Radiobutton(self, text="動作追蹤系統", value="0", variable=self.method,font=('微軟正黑體',10))
        method1.place(x=120,y=210)
        method2 = tk.Radiobutton(self, text="深度相機 *", value="1", variable= self.method,font=('微軟正黑體',10))
        method2.place(x=355,y=210)
        #method3 = tk.Radiobutton(self, text="45°+180°深度相機", value="2", variable= self.method,font=('微軟正黑體',10))
        #method3.place(x=205,y=360)
        method4 = tk.Radiobutton(self, text="深度相機", value="3", variable= self.method,font=('微軟正黑體',10))
        method4.place(x=575,y=210)
        photo = tk.PhotoImage(file = "img/right.png")
        photoimage = photo.subsample(6, 6) 
        photo2 = tk.PhotoImage(file = "img/left.png")
        photoimage2 = photo2.subsample(6, 6) 
        button1 = tk.Button(self,image = photoimage,command=lambda: [get_data()],padx=5,pady=5)
        button1.image = photoimage
        button2 = tk.Button(self, image = photoimage2,command=lambda: [controller.show_frame("PageTwo")],padx=5,pady=5)
        button2.image = photoimage2
        button1.place(x=757,y=150)
        button2.place(x=0,y=150) 
      
class PageFour1(tk.Frame): 
    def correct_label(self):
        if self.controller.shared_data["sex"] == '1' or self.controller.shared_data["sex"] == '男性':        
            self.controller.shared_data["sex"] = "男性"
        else:
            self.controller.shared_data["sex"] = "女性"
        self.label2.config(text="工作人員: "+str(self.controller.shared_data["username"]))
        self.label3.config(text="年齡: "+str(self.controller.shared_data["year"]))
        self.label4.config(text="性別: "+str(self.controller.shared_data["sex"]))
        self.label5.config(text="身高: "+str(self.controller.shared_data["height"])+"cm")
        self.label6.config(text="體重: "+str(self.controller.shared_data["weight"])+"kg")
        self.label7.config(text="搬運姿勢: "+str(self.controller.shared_data["origin"])+" 到 "+str(self.controller.shared_data["end"]))
        self.label8.config(text="箱子重量: "+str(self.controller.shared_data["box"])+"kg")
        self.label9.config(text="搬運日期: "+str(self.controller.shared_data["date"]))
  
    def __init__(self, parent, controller): 
        def open_file(): 
            file = fd.askopenfilename()
            self.label12.configure(image = photo4)
            self.label12.image = photo4
            self.label12.place(x=500,y=60)
            text.set(file)          
        def check_data():             
            res = not all(self.controller.shared_data.values())
            if res == True or len(text.get()) == 0:
                messagebox.showerror("Error", "資料不齊全，請確認!")
            else:
                controller.show_frame("PageProcess")
                self.controller.frames["PageProcess"].progress2(text.get())

                           
        tk.Frame.__init__(self, parent)        
        self.controller = controller        
        label = tk.Label(self, text="資料上傳",font='微軟正黑體 14 bold')        
        label.pack(pady=2)    
        photo3 = tk.PhotoImage(file = "img/not_upload.gif")
        photo3 = photo3.subsample(3, 3)
        photo4 = tk.PhotoImage(file = "img/uploaded.png")
        photo4 = photo4.subsample(3, 3) 
        self.label12 = tk.Label(self,image = photo3)
        self.label12.image = photo3
        self.label12.place(x=510,y=60)
        frame = tk.LabelFrame(self, text = "檔案上傳區", padx=5,pady=5, labelanchor="n",font='微軟正黑體 12 bold')
        frame.place(x=400,y=150)
        label10 = tk.Label(frame,text="深度相機檔案",width=40)
        label10.pack()
        btn = tk.Button(frame, text ='開啟檔案', command = lambda:open_file()) 
        btn.pack(side = 'top', pady = 5)
        text = tk.StringVar()
        text.set("")
        label11 = tk.Label(frame,textvariable = text,width=40)
        label11.pack()
        self.label2 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label2.place(x=150,y=50) 
        self.label3 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label3.place(x=150,y=90) 
        self.label4 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label4.place(x=150,y=130)
        self.label5 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label5.place(x=150,y=170) 
        self.label6 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label6.place(x=150,y=210) 
        self.label7 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label7.place(x=150,y=250) 
        self.label8 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label8.place(x=150,y=290) 
        self.label9 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label9.place(x=150,y=330)                
        photo = tk.PhotoImage(file = "img/right.png")
        photoimage = photo.subsample(6, 6) 
        photo2 = tk.PhotoImage(file = "img/left.png")
        photoimage2 = photo2.subsample(6, 6) 
        button1 = tk.Button(self,image = photoimage,command= lambda: [check_data()],padx=5,pady=5)
        button1.image = photoimage
        button2 = tk.Button(self, image = photoimage2,command=lambda: [controller.show_frame("PageThree")],padx=5,pady=5)
        button2.image = photoimage2
        button1.place(x=757,y=150)
        button2.place(x=0,y=150)     

class PageProcess(tk.Frame):
    def progress(self,currentValue,data_file):
        self.progressbar["value"]=currentValue 
        self.label3.configure(text = str(currentValue)+"%")
        if currentValue == 100: 
            if self.controller.shared_data["method"] == '0':                    
                k = idm.get_data(self.controller.shared_data, data_file)
                self.controller.frames["PageFive1"].draw(k)
                self.controller.show_frame("PageFive1")
                
            elif self.controller.shared_data["method"] == '1': 
                k = k45.get_data(self.controller.shared_data, data_file, kinect1)
                self.controller.show_frame("PageFive2")
                self.controller.frames["PageFive2"].draw(k) 
            elif self.controller.shared_data["method"] == '2':        
                k = k45.get_data2(self.controller.shared_data, data_file, kinect1)
                self.controller.show_frame("PageFive2")
                self.controller.frames["PageFive2"].draw(k) 
            else:
                k = k45.get_data3(self.controller.shared_data, data_file, kinect1)
                self.controller.show_frame("PageFive2")
                self.controller.frames["PageFive2"].draw(k)          
    
    def progress2(self,data_file):  
        maxValue=100
        currentValue=0
        self.progressbar["value"]=currentValue
        self.progressbar["maximum"]=maxValue
        divisions=10
        for i in range(divisions):
            currentValue=currentValue+10           
            self.progressbar.after(500,PageProcess.progress(self,currentValue,data_file))
            self.progressbar.update()      

    def __init__(self, parent, controller):         
        tk.Frame.__init__(self, parent)
        self.controller = controller        
        label = tk.Label(self,text = "資料處理" ,font='微軟正黑體 14 bold')
        label.pack(pady=2)
        label2 = tk.Label(self,text = "下背負荷計算中，請稍後..." ,font='微軟正黑體 10')
        label2.place(x=320,y=180)        
        self.label3 = tk.Label(self,text = "0" ,font='微軟正黑體 10 bold')
        self.label3.place(x=380,y=120)  
        self.progressbar=ttk.Progressbar(self,orient="horizontal",length=300,mode="determinate")
        self.progressbar.place(x=250,y=150) 

class PageFive1(tk.Frame):
    def draw(self, data):        
        self.label2.config(text="工作人員: "+str(self.controller.shared_data["username"]))        
        self.label3.config(text="搬運姿勢: "+str(self.controller.shared_data["origin"])+" 到 "+str(self.controller.shared_data["end"]))
        self.label4.config(text="箱子重量: "+str(self.controller.shared_data["box"])+"kg")
        self.label5.config(text="搬運日期: "+str(self.controller.shared_data["date"]))    
        self.label6.config(text="平均下背力矩: "+str(abs(round(float(statistics.mean(np.sum(data[0][1],axis=1))),2)))+" Nm")
        self.label7.config(text="最大下背力矩: "+str(abs(round(float(min(np.sum(data[0][1],axis=1))),2)))+" Nm")
        self.label8.config(text="平均下背壓力: "+str(round(float(np.mean(data[1])),2))+" N")
        self.label9.config(text="最大下背壓力: "+str(round(float(np.amax(data[1])),2))+" N")    
        l5s1_data.append(data)        
        joint_data.append(data[2])
        PageFive1.compressvie_force(self,2)
        
    def compressvie_force(self, num): 
        global ani
        class Arrow3D(FancyArrowPatch):
            def __init__(self, xs, ys, zs, *args, **kwargs):
                FancyArrowPatch.__init__(self, (0,0), (0,0), *args, **kwargs)
                self._verts3d = xs, ys, zs
            
            def draw(self, renderer):
                xs3d, ys3d, zs3d = self._verts3d
                xs, ys, zs = proj3d.proj_transform(xs3d, ys3d, zs3d, renderer.M)
                self.set_positions((xs[0],ys[0]),(xs[1],ys[1]))
                FancyArrowPatch.draw(self, renderer)  
        def update_graph(num):                    
            ax.lines = []                
            graph._offsets3d = (joint_data[0][:,0::3][num],joint_data[0][:,1::3][num],joint_data[0][:,2::3][num])
            for i in joint_plot2:
                line = plt3d.art3d.Line3D(joint_data[0][:,i][num],joint_data[0][:,[i[0]+1,i[1]+1]][num],joint_data[0][:,[i[0]+2,i[1]+2]][num])
                ax.add_line(line)
            return graph, line
        if num == 0:
            self.f.clf()
            self.a = self.f.add_subplot(111) 
            self.a.clear()
            data, = self.a.plot(l5s1_data[0][1], label="Compressive Force") 
            self.a.set_title('L5S1 Compressive Force',fontsize='xx-small')
            self.a.set_xlabel('Frame',fontsize='xx-small')
            self.a.set_ylabel('Force (N)',fontsize='xx-small')  
            self.a.tick_params(axis = 'both',labelsize=8)
            self.a.legend(handles=[data], loc='upper right',fontsize="xx-small")
            self.canvas.draw()
        elif num == 1: 
            self.f.clf()
            self.a = self.f.add_subplot(111) 
            self.a.clear()
            mx, = self.a.plot(l5s1_data[0][0][1][:,0], label = "Mx")
            my, = self.a.plot(l5s1_data[0][0][1][:,1], label = "My") 
            mz, = self.a.plot(l5s1_data[0][0][1][:,2], label = "Mz")             
            self.a.set_title('L5S1 Moment',fontsize='xx-small')
            self.a.set_xlabel('Frame',fontsize='xx-small')
            self.a.set_ylabel('Force (Nm)',fontsize='xx-small')  
            self.a.tick_params(axis = 'both',labelsize=8)
            self.a.legend(handles=[mx, my, mz], loc='center right', fontsize="xx-small")
            self.canvas.draw()
                                  
        elif num == 2:              
            ax = self.f.add_subplot(111, projection='3d')
            ax.set_xlabel('x')
            ax.set_ylabel('y')
            ax.set_zlabel('z')
            ax.set_xlim(-1, 1)
            ax.set_ylim(-1, 1)
            ax.set_zlim(-1, 2)
            arrow_prop_dict = dict(mutation_scale=20, arrowstyle='->',shrinkA=0, shrinkB=0)
            a = Arrow3D([-0.75,-0.25], [-0.75,-0.75], [-0.75,-0.75], **arrow_prop_dict, color='r')
            ax.add_artist(a)
            a = Arrow3D([-0.75,-0.75], [-0.75,-0.25], [-0.75,-0.75], **arrow_prop_dict, color='b')
            ax.add_artist(a)
            a = Arrow3D([-0.75,-0.75], [-0.75,-0.75], [-0.75,0.25], **arrow_prop_dict, color='g')
            ax.add_artist(a)
            ax.text(0, -1, -1, r'$x$')
            ax.text(-0.75, -0.25, -1, r'$y$')
            ax.text(-0.5, -1, 0, r'$z$')
            graph = ax.scatter([],[],[],color="red")            
            line = ax.plot([],[],[],color = "blue")
            ax.view_init(elev=45, azim=135)
            ani = animation.FuncAnimation(self.f, update_graph,frames = joint_data[0].shape[0]-4, blit=False, interval = 20)
            self.canvas.draw()            
        elif num == 3:  
            def save():
                filename =  fd.asksaveasfilename(initialdir = "/",title = "Select file",filetypes=[('excel file','*.xlsx')], defaultextension = '.xlsx')                
                idm.save_data(filename)
                #idm.save_gif(filename, ani)
                tk.messagebox.showinfo(title='Notify', message='數據匯出成功!!')
            save() 
    
        
    def __init__(self, parent, controller):       
        tk.Frame.__init__(self, parent)
        self.controller = controller 
        label = tk.Label(self, text = "評估結果" ,font='微軟正黑體 14 bold')
        label.pack()
        frame = tk.LabelFrame(self, text = "下背負荷", padx=5,pady=5, labelanchor="n",font='微軟正黑體 12 bold')
        frame.place(x=50,y=50)
        self.label2 = tk.Label(frame, text = "",font='微軟正黑體 10')
        self.label2.pack(pady = 5)
        self.label3 = tk.Label(frame, text ="",font='微軟正黑體 10')
        self.label3.pack(pady = 5)
        self.label4 = tk.Label(frame, text ="",font='微軟正黑體 10')
        self.label4.pack(pady = 5)
        self.label5 = tk.Label(frame, text ="",font='微軟正黑體 10')
        self.label5.pack(pady = 5)  
        self.label10 = tk.Label(frame, text ="----------",font='微軟正黑體 10')
        self.label10.pack(pady = 5)
        self.label6 = tk.Label(frame, text ="",font='微軟正黑體 10')
        self.label6.pack(pady = 5)
        self.label7 = tk.Label(frame, text ="",font='微軟正黑體 10')
        self.label7.pack(pady = 5)
        self.label8 = tk.Label(frame, text ="",font='微軟正黑體 10')
        self.label8.pack(pady = 5)
        self.label9 = tk.Label(frame, text ="",font='微軟正黑體 10')
        self.label9.pack(pady = 5)
        self.btn4 = tk.Button(self,text = "關節骨架" ,command= lambda: [PageFive1.compressvie_force(self, 2)],padx=5,pady=5)
        self.btn4.place(x=285,y=150)
        self.btn = tk.Button(self,text = "下背壓力" ,command= lambda: [PageFive1.compressvie_force(self, 0)],padx=5,pady=5)
        self.btn.place(x=285,y=200)
        self.btn2 = tk.Button(self,text = "下背力矩" ,command= lambda: [PageFive1.compressvie_force(self, 1)],padx=5,pady=5)
        self.btn2.place(x=285,y=250)
        self.btn3 = tk.Button(self,text = "數據匯出" ,command= lambda: [PageFive1.compressvie_force(self, 3)],padx=5,pady=5)
        self.btn3.place(x=285,y=300)
        self.f = Figure(figsize=(4.5,3)) 
        self.a = self.f.add_subplot(111)
        self.cv = Canvas(self,height=100, width=100)
        self.cv.place(x=360,y=50)
        self.canvas = FigureCanvasTkAgg(self.f, self.cv) 
        self.canvas.draw()
        self.canvas.get_tk_widget().pack()
        toolbar = NavigationToolbar2Tk(self.canvas, self.cv)
        toolbar.update()
        self.canvas._tkcanvas.pack(expand=1)

class PageFive2(tk.Frame):
    def draw(self, data):             
        self.label2.config(text="工作人員: "+str(self.controller.shared_data["username"]))        
        self.label3.config(text="搬運姿勢: "+str(self.controller.shared_data["origin"])+" 到 "+str(self.controller.shared_data["end"]))
        self.label4.config(text="箱子重量: "+str(self.controller.shared_data["box"])+"kg")
        self.label5.config(text="搬運日期: "+str(self.controller.shared_data["date"]))    
        self.label6.config(text="平均下背力矩: "+str(abs(round(float(statistics.mean(np.sum(data[0],axis=1))),2)))+" Nm")
        self.label7.config(text="最大下背力矩: "+str(abs(round(float(min(np.sum(data[0],axis=1))),2)))+" Nm")
        self.label8.config(text="平均下背壓力: "+str(round(float(np.mean(data[1])),2))+" N")
        self.label9.config(text="最大下背壓力: "+str(round(float(np.amax(data[1])),2))+" N")    
        l5s1_data.append(data)
        joint_data.append(data[2])
        PageFive2.compressvie_force(self,2)        
        
    def compressvie_force(self, num): 
        global ani
        class Arrow3D(FancyArrowPatch):
            def __init__(self, xs, ys, zs, *args, **kwargs):
                FancyArrowPatch.__init__(self, (0,0), (0,0), *args, **kwargs)
                self._verts3d = xs, ys, zs
        
            def draw(self, renderer):
                xs3d, ys3d, zs3d = self._verts3d
                xs, ys, zs = proj3d.proj_transform(xs3d, ys3d, zs3d, renderer.M)
                self.set_positions((xs[0],ys[0]),(xs[1],ys[1]))
                FancyArrowPatch.draw(self, renderer)  
        def update_graph(num):
            num += 2    
            ax.lines = []       
            graph._offsets3d = (joint_data[0][:,1::3][num],joint_data[0][:,2::3][num],joint_data[0][:,3::3][num])
            for i in joint_plot:
                line = plt3d.art3d.Line3D(joint_data[0][:,i][num],joint_data[0][:,[i[0]+1,i[1]+1]][num],joint_data[0][:,[i[0]+2,i[1]+2]][num])
                ax.add_line(line)
            return graph, line
        if num == 0:            
            self.f.clf()
            self.a = self.f.add_subplot(111) 
            self.a.clear()
            data, = self.a.plot(l5s1_data[0][1], label="Compressive Force") 
            self.a.set_title('L5S1 Compressive Force',fontsize='xx-small')
            self.a.set_xlabel('Frame',fontsize='xx-small')
            self.a.set_ylabel('Force (N)',fontsize='xx-small')  
            self.a.tick_params(axis = 'both',labelsize=8)
            self.a.legend(handles=[data], loc='upper right',fontsize="xx-small")
            self.canvas.draw()
        elif num == 1: 
            self.f.clf()
            self.a = self.f.add_subplot(111) 
            self.a.clear()
            mx, = self.a.plot(l5s1_data[0][0][:,0], label = "Mx")
            my, = self.a.plot(l5s1_data[0][0][:,1], label = "My") 
            mz, = self.a.plot(l5s1_data[0][0][:,2], label = "Mz")             
            self.a.set_title('L5S1 Moment',fontsize='xx-small')
            self.a.set_xlabel('Frame',fontsize='xx-small')
            self.a.set_ylabel('Force (Nm)',fontsize='xx-small')  
            self.a.tick_params(axis = 'both',labelsize=8)
            self.a.legend(handles=[mx, my, mz], loc='center right', fontsize="xx-small")
            self.canvas.draw()            
        elif num == 2:              
            ax = self.f.add_subplot(111, projection='3d')
            ax.set_xlabel('x')
            ax.set_ylabel('y')
            ax.set_zlabel('z')
            ax.set_xlim(-1, 1)
            ax.set_ylim(-1, 1)
            ax.set_zlim(-1, 2)
            arrow_prop_dict = dict(mutation_scale=20, arrowstyle='->',shrinkA=0, shrinkB=0)
            a = Arrow3D([-0.75,-0.25], [-0.75,-0.75], [-0.75,-0.75], **arrow_prop_dict, color='r')
            ax.add_artist(a)
            a = Arrow3D([-0.75,-0.75], [-0.75,-0.25], [-0.75,-0.75], **arrow_prop_dict, color='b')
            ax.add_artist(a)
            a = Arrow3D([-0.75,-0.75], [-0.75,-0.75], [-0.75,0.25], **arrow_prop_dict, color='g')
            ax.add_artist(a)
            ax.text(0, -1, -1, r'$x$')
            ax.text(-0.75, -0.25, -1, r'$y$')
            ax.text(-0.5, -1, 0, r'$z$')
            graph = ax.scatter([],[],[],color="red")            
            line = ax.plot([],[],[],color = "blue")
            ax.view_init(elev=45, azim=135)
            ani = animation.FuncAnimation(self.f, update_graph,frames = joint_data[0].shape[0]-4, blit=False, interval = 20)
            self.canvas.draw()            
        elif num == 3:            
            def save():                
                filename =  fd.asksaveasfilename(initialdir = "/",title = "Select file",filetypes=[('excel file','*.xlsx')], defaultextension = '.xlsx')                
                k45.save_data(filename)
                #k45.save_gif(filename, ani)
                tk.messagebox.showinfo(title='Notify', message='數據匯出成功!!')
            save()        
    def __init__(self, parent, controller):       
        tk.Frame.__init__(self, parent)
        self.controller = controller 
        frame = tk.LabelFrame(self, text = "下背負荷", padx=5,pady=5, labelanchor="n",font='微軟正黑體 12 bold')
        frame.place(x=50,y=50)
        self.label2 = tk.Label(frame, text = "",font='微軟正黑體 10')
        self.label2.pack(pady = 5)
        self.label3 = tk.Label(frame, text ="",font='微軟正黑體 10')
        self.label3.pack(pady = 5)
        self.label4 = tk.Label(frame, text ="",font='微軟正黑體 10')
        self.label4.pack(pady = 5)
        self.label5 = tk.Label(frame, text ="",font='微軟正黑體 10')
        self.label5.pack(pady = 5)  
        self.label10 = tk.Label(frame, text ="----------",font='微軟正黑體 10')
        self.label10.pack(pady = 5)
        self.label6 = tk.Label(frame, text ="",font='微軟正黑體 10')
        self.label6.pack(pady = 5)
        self.label7 = tk.Label(frame, text ="",font='微軟正黑體 10')
        self.label7.pack(pady = 5)
        self.label8 = tk.Label(frame, text ="",font='微軟正黑體 10')
        self.label8.pack(pady = 5)
        self.label9 = tk.Label(frame, text ="",font='微軟正黑體 10')
        self.label9.pack(pady = 5)
        self.btn4 = tk.Button(self,text = "關節骨架" ,command= lambda: [PageFive2.compressvie_force(self, 2)],padx=5,pady=5)
        self.btn4.place(x=285,y=150)
        self.btn = tk.Button(self,text = "下背壓力" ,command= lambda: [PageFive2.compressvie_force(self, 0)],padx=5,pady=5)
        self.btn.place(x=285,y=200)
        self.btn2 = tk.Button(self,text = "下背力矩" ,command= lambda: [PageFive2.compressvie_force(self, 1)],padx=5,pady=5)
        self.btn2.place(x=285,y=250)
        self.btn3 = tk.Button(self,text = "數據匯出" ,command= lambda: [PageFive2.compressvie_force(self, 3)],padx=5,pady=5)
        self.btn3.place(x=285,y=300)
        self.f = Figure(figsize=(4.5,3)) 
        self.a = self.f.add_subplot(111)
        self.cv = Canvas(self,height=100, width=100)
        self.cv.place(x=360,y=50)
        self.canvas = FigureCanvasTkAgg(self.f, self.cv)  
        self.canvas.draw()
        self.canvas.get_tk_widget().pack()
        toolbar = NavigationToolbar2Tk(self.canvas, self.cv)
        toolbar.update()
        self.canvas._tkcanvas.pack(expand=1)
                
class PageFour2(tk.Frame): 
    
    def correct_label(self):
        if self.controller.shared_data["sex"] == '1' or self.controller.shared_data["sex"] == '男性':        
            self.controller.shared_data["sex"] = "男性"
        else:
            self.controller.shared_data["sex"] = "女性"
        self.label2.config(text="工作人員: "+str(self.controller.shared_data["username"]))
        self.label3.config(text="年齡: "+str(self.controller.shared_data["year"]))
        self.label4.config(text="性別: "+str(self.controller.shared_data["sex"]))
        self.label5.config(text="身高: "+str(self.controller.shared_data["height"])+"cm")
        self.label6.config(text="體重: "+str(self.controller.shared_data["weight"])+"kg")
        self.label7.config(text="搬運姿勢: "+str(self.controller.shared_data["origin"])+" 到 "+str(self.controller.shared_data["end"]))
        self.label8.config(text="箱子重量: "+str(self.controller.shared_data["box"])+"kg")
        self.label9.config(text="搬運日期: "+str(self.controller.shared_data["date"]))
  
    def __init__(self, parent, controller): 
        def open_file(): 
            file = fd.askopenfilename()
            self.label12.configure(image = photo4)
            self.label12.image = photo4
            self.label12.place(x=500,y=40)
            text.set(file)          
        def check_data():             
            res = not all(self.controller.shared_data.values())
            if timeString.get() != "" or timeString2.get() != "":
                kinect1.append(timeString.get())
                kinect1.append(timeString2.get())            

            if res == True or len(text.get()) == 0 or len(kinect1) == 0:
                messagebox.showerror("Error", "資料不齊全，請確認!")
            else:
                controller.show_frame("PageProcess")
                self.controller.frames["PageProcess"].progress2(text.get())                           
        tk.Frame.__init__(self, parent)        
        self.controller = controller        
        label = tk.Label(self, text="資料上傳",font='微軟正黑體 14 bold')        
        label.pack(pady=2)    
        photo3 = tk.PhotoImage(file = "img/not_upload.gif")
        photo3 = photo3.subsample(3, 3)
        photo4 = tk.PhotoImage(file = "img/uploaded.png")
        photo4 = photo4.subsample(3, 3) 
        self.label12 = tk.Label(self,image = photo3)
        self.label12.image = photo3
        self.label12.place(x=510,y=40)
        frame = tk.LabelFrame(self, text = "檔案上傳區", padx=5,pady=5, labelanchor="n",font='微軟正黑體 12 bold')
        frame.place(x=400,y=130)
        label10 = tk.Label(frame,text="深度相機檔案",width=40)
        label10.pack()
        btn = tk.Button(frame, text ='開啟檔案', command = lambda:open_file()) 
        btn.pack(side = 'top', pady = 5)
        text = tk.StringVar()
        text.set("")
        label11 = tk.Label(frame,textvariable = text,width=40)
        label11.pack()
        self.label2 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label2.place(x=150,y=50) 
        self.label3 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label3.place(x=150,y=90) 
        self.label4 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label4.place(x=150,y=130)
        self.label5 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label5.place(x=150,y=170) 
        self.label6 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label6.place(x=150,y=210) 
        self.label7 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label7.place(x=150,y=250) 
        self.label8 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label8.place(x=150,y=290) 
        self.label9 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label9.place(x=150,y=330)         
        label12 = tk.Label(frame, text="Time1:")
        label12.pack()
        timeString = tk.StringVar()
        self.entry = tk.Entry(frame, textvariable = timeString)
        self.entry.pack()
        label13 = tk.Label(frame, text="Time2:")
        label13.pack()
        timeString2 = tk.StringVar()
        self.entry2 = tk.Entry(frame, textvariable = timeString2)
        self.entry2.pack()
        photo = tk.PhotoImage(file = "img/right.png")
        photoimage = photo.subsample(6, 6) 
        photo2 = tk.PhotoImage(file = "img/left.png")
        photoimage2 = photo2.subsample(6, 6) 
        button1 = tk.Button(self,image = photoimage,command= lambda: [check_data()],padx=5,pady=5)
        button1.image = photoimage
        button2 = tk.Button(self, image = photoimage2,command=lambda: [controller.show_frame("PageThree")],padx=5,pady=5)
        button2.image = photoimage2
        button1.place(x=757,y=150)
        button2.place(x=0,y=150) 

class PageFour3(tk.Frame): 
    
    def correct_label(self):
        if self.controller.shared_data["sex"] == '1' or self.controller.shared_data["sex"] == '男性':        
            self.controller.shared_data["sex"] = "男性"
        else:
            self.controller.shared_data["sex"] = "女性"
        self.label2.config(text="工作人員: "+str(self.controller.shared_data["username"]))
        self.label3.config(text="年齡: "+str(self.controller.shared_data["year"]))
        self.label4.config(text="性別: "+str(self.controller.shared_data["sex"]))
        self.label5.config(text="身高: "+str(self.controller.shared_data["height"])+"cm")
        self.label6.config(text="體重: "+str(self.controller.shared_data["weight"])+"kg")
        self.label7.config(text="搬運姿勢: "+str(self.controller.shared_data["origin"])+" 到 "+str(self.controller.shared_data["end"]))
        self.label8.config(text="箱子重量: "+str(self.controller.shared_data["box"])+"kg")
        self.label9.config(text="搬運日期: "+str(self.controller.shared_data["date"]))
  
    def __init__(self, parent, controller): 
        def open_file(num):
            if num == 0:
                file = fd.askopenfilename()                
                self.label12.configure(image = photo4)
                self.label12.image = photo4
                self.label12.place(x=320,y=40)
                text.set(file)   
            elif num == 1:
                file2 = fd.askopenfilename() 
                self.label24.configure(image = photo4)
                self.label24.image = photo4
                self.label24.place(x=570,y=40)
                text2.set(file2) 
        def check_data():             
            res = not all(self.controller.shared_data.values())
            if timeString.get() != "" or timeString2.get() != "":
                
                kinect1.append(timeString.get()+'#'+timeString3.get())
                kinect1.append(timeString2.get()+'#'+timeString4.get()) 
            if res == True or len(text.get()) == 0 or len(kinect1) == 0 or len(text2.get()) == 0:
                messagebox.showerror("Error", "資料不齊全，請確認!")
            else:
                k = text.get() + '#' + text2.get()                
                controller.show_frame("PageProcess")
                self.controller.frames["PageProcess"].progress2(k)                           
        tk.Frame.__init__(self, parent)  
        self.controller = controller        
        label = tk.Label(self, text="資料上傳",font='微軟正黑體 14 bold')        
        label.pack(pady=2)    
        photo3 = tk.PhotoImage(file = "img/not_upload.gif")
        photo3 = photo3.subsample(3, 3)
        photo4 = tk.PhotoImage(file = "img/uploaded.png")
        photo4 = photo4.subsample(3, 3) 
        self.label12 = tk.Label(self,image = photo3)
        self.label12.image = photo3
        self.label12.place(x=320,y=40)
        self.label24 = tk.Label(self,image = photo3)
        self.label24.image = photo3
        self.label24.place(x=570,y=40)
        frame = tk.LabelFrame(self, text = "檔案上傳區", padx=5,pady=5, labelanchor="n",font='微軟正黑體 12 bold')
        frame.place(x=250,y=130)
        frame2 = tk.LabelFrame(self, text = "檔案上傳區", padx=5,pady=5, labelanchor="n",font='微軟正黑體 12 bold')
        frame2.place(x=500,y=130)
        label10 = tk.Label(frame,text="深度相機檔案",width=30)
        label10.pack()
        label20 = tk.Label(frame2,text="深度相機檔案",width=30)
        label20.pack()
        text = tk.StringVar()
        text.set("")
        btn = tk.Button(frame, text ='開啟檔案', command = lambda:open_file(0)) 
        btn.pack(side = 'top', pady = 2)
        btn2 = tk.Button(frame2, text ='開啟檔案', command = lambda:open_file(1)) 
        btn2.pack(side = 'top', pady = 2)
        label11 = tk.Label(frame,textvariable = text,width=30)
        label11.pack()
        text2 = tk.StringVar()
        text2.set("")
        label21 = tk.Label(frame2,textvariable = text2)
        label21.pack()
        label22 = tk.Label(frame2, text="Time1:")
        label22.pack()
        timeString3 = tk.StringVar()
        self.entry3 = tk.Entry(frame2, textvariable = timeString3)
        self.entry3.pack()
        label23 = tk.Label(frame2, text="Time2:")
        label23.pack()
        timeString4 = tk.StringVar()
        self.entry4 = tk.Entry(frame2, textvariable = timeString4)
        self.entry4.pack()
        self.label2 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label2.place(x=80,y=80) 
        self.label3 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label3.place(x=80,y=120) 
        self.label4 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label4.place(x=80,y=160)
        self.label5 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label5.place(x=80,y=200) 
        self.label6 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label6.place(x=80,y=240) 
        self.label7 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label7.place(x=80,y=280) 
        self.label8 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label8.place(x=80,y=320) 
        self.label9 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label9.place(x=80,y=360)         
        label12 = tk.Label(frame, text="Time1:")
        label12.pack()
        timeString = tk.StringVar()
        self.entry = tk.Entry(frame, textvariable = timeString)
        self.entry.pack()
        label13 = tk.Label(frame, text="Time2:")
        label13.pack()
        timeString2 = tk.StringVar()
        self.entry2 = tk.Entry(frame, textvariable = timeString2)
        self.entry2.pack()
        photo = tk.PhotoImage(file = "img/right.png")
        photoimage = photo.subsample(6, 6) 
        photo2 = tk.PhotoImage(file = "img/left.png")
        photoimage2 = photo2.subsample(6, 6) 
        button1 = tk.Button(self,image = photoimage,command= lambda: [check_data()],padx=5,pady=5)
        button1.image = photoimage
        button2 = tk.Button(self, image = photoimage2,command=lambda: [controller.show_frame("PageThree")],padx=5,pady=5)
        button2.image = photoimage2
        button1.place(x=757,y=150)
        button2.place(x=0,y=150) 
        
class PageFour4(tk.Frame): 
    
    def correct_label(self):
        if self.controller.shared_data["sex"] == '1' or self.controller.shared_data["sex"] == '男性':        
            self.controller.shared_data["sex"] = "男性"
        else:
            self.controller.shared_data["sex"] = "女性"
        self.label2.config(text="工作人員: "+str(self.controller.shared_data["username"]))
        self.label3.config(text="年齡: "+str(self.controller.shared_data["year"]))
        self.label4.config(text="性別: "+str(self.controller.shared_data["sex"]))
        self.label5.config(text="身高: "+str(self.controller.shared_data["height"])+"cm")
        self.label6.config(text="體重: "+str(self.controller.shared_data["weight"])+"kg")
        self.label7.config(text="搬運姿勢: "+str(self.controller.shared_data["origin"])+" 到 "+str(self.controller.shared_data["end"]))
        self.label8.config(text="箱子重量: "+str(self.controller.shared_data["box"])+"kg")
        self.label9.config(text="搬運日期: "+str(self.controller.shared_data["date"]))
  
    def __init__(self, parent, controller): 
        def open_file(num):
            if num == 1:
                file2 = fd.askopenfilename() 
                label11.configure(image = photo4)
                label11.image = photo4
                label11.place(x=500,y=60)
                text2.set(file2)             

        def check_data():             
            res = not all(self.controller.shared_data.values())            
            if res == True or  len(text2.get()) ==0:
                messagebox.showerror("Error", "資料不齊全，請確認!")
            else:
                k =  text2.get()                
                controller.show_frame("PageProcess")
                self.controller.frames["PageProcess"].progress2(k)    
                       
        tk.Frame.__init__(self, parent)  
        self.controller = controller 
        photo3 = tk.PhotoImage(file = "img/not_upload.gif")
        photo3 = photo3.subsample(3, 3)
        photo4 = tk.PhotoImage(file = "img/uploaded.png")
        photo4 = photo4.subsample(3, 3) 
        text2 = tk.StringVar()
        text2.set("")

        label = tk.Label(self, text="資料上傳",font='微軟正黑體 14 bold')        
        label.pack(pady=2)
        self.label2 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label2.place(x=150,y=50) 
        self.label3 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label3.place(x=150,y=90) 
        self.label4 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label4.place(x=150,y=130)
        self.label5 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label5.place(x=150,y=170) 
        self.label6 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label6.place(x=150,y=210) 
        self.label7 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label7.place(x=150,y=250) 
        self.label8 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label8.place(x=150,y=290) 
        self.label9 = tk.Label(self, text ="",font='微軟正黑體 10 bold')
        self.label9.place(x=150,y=330)  
        

        label11 = tk.Label(self,image = photo3)
        label11.image = photo3
        label11.place(x=510,y=60)
      
        
        frame2 = tk.LabelFrame(self, text = "檔案上傳區", padx=5,pady=5, labelanchor="n",font='微軟正黑體 12 bold')
        frame2.place(x=400,y=150)
        label17 = tk.Label(frame2, text="深度相機檔案", width = 40)
        label17.pack()
        label18 = tk.Label(frame2, textvariable = text2,width = 40)
        label18.pack()
        btn3 = tk.Button(frame2, text ='開啟檔案', command = lambda:open_file(1)) 
        btn3.pack(side = 'top', pady = 2)       
        


        photo = tk.PhotoImage(file = "img/right.png")
        photoimage = photo.subsample(6, 6) 
        photo2 = tk.PhotoImage(file = "img/left.png")
        photoimage2 = photo2.subsample(6, 6) 
        button1 = tk.Button(self,image = photoimage,command= lambda: [check_data()],padx=5,pady=5)
        button1.image = photoimage
        button2 = tk.Button(self, image = photoimage2,command=lambda: [controller.show_frame("PageThree")],padx=5,pady=5)
        button2.image = photoimage2
        button1.place(x=757,y=150)
        button2.place(x=0,y=150) 
if __name__ == "__main__":    
    app = SampleApp()    
    app.geometry("800x400")
    app.resizable(False, False)
    app.mainloop()