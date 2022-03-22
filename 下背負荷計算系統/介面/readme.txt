<<資料夾>>
img資料夾: 介面圖庫
mp4:介面gif檔
properties: 男女肢段重量、關節中心位置資訊 

<<py檔>>
low_back_interface: 下背負荷介面 (主要執行程式)
Inverse_dynamic_model: motion下背負荷計算code (bottom-up + top-down)
kinect: 深度相機下背負荷計算code (top-down)

※執行前安裝的套件 (搜尋開啟anaconda prompt輸入以下指令安裝套件)
1. pip install tkcalendar
2. pip install imageio-ffmpeg

※若執行low_back_interface出現'PhotoImage' object has no attribute '_PhotoImage__photo'
1.關掉介面，重新執行程式

2.若仍無法執行程式，註解以下code (在 class StartPage下、程式第80-87行)
video_name = "mp4/icon.mp4" 
video = imageio.get_reader(video_name)
label = tk.Label(self,width=240, height=240)
label.config(bg="white")
label.pack(pady=20)
thread = threading.Thread(target=stream, args=(label,))
thread.daemon = 1
thread.start()  

※若執行low_back_interface出現其他error 如'pyimage doesn't exist'
到spyder上面選擇Consoles，選restart kernel，再重新執行程式


※Kinect數據濾波
原始數據濾波line 163: data2[:,i] = low_pass(data2[:,i],3, 125,4) 4th order Butterworth low pass filter 3Hz
運動學數據濾波 line 74: vel = low_pass(vel,8,125,4)
               line 77: acc = low_pass(acc,8,125,4)

