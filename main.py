from functools import partial
from tkinter import messagebox
import datetime
from win32com.client import Dispatch
from tkinter import *
import PIL.Image, PIL.ImageTk
from tkinter import *

'''
    Author: Aryan
    Date: 19 April 2022
    Purpose: Making Personal Time Manager
'''

#------------------------------------------------------------------

def average():
  # Return total time worked in per day
  ave_set=[]
  f=open("Timmer.txt")
  lines = f.readlines()
  for i1 in lines:
    b=1
    i1_s = i1[0:10]
    for i2 in lines:
      i2_s = i2[0:10]
      if i1_s == i2_s:
        lis1 = i1.split("-->")
        lis2=lis1[1].split("and")
        i1_s_min = float(lis2[0].replace("min",""))
        ave_set.append(f"{i1_s}={i1_s_min}")
      b+=1
  ave_set=set(ave_set).union()
  ave_dict={}
  for item in ave_set:
    item1=item.split("=")
    if item1[0] in ave_dict:
      min1 = float(item1[1])
      min1+=float(ave_dict[item1[0]])
      ave_dict[item1[0]]=min1
    else:
      ave_dict[item1[0]]=float(item1[1] )
    f.close()
    min1=0
  return ave_dict

def month():
  '''
  first it give dict of month:(time,time,..) second it give dict of month name and total time
  '''
  months =  ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
  ave_dict= average()
  x=1
  total=0
  all_info={}
  Total_time = {}
  times = []
  y=0
  for month_name in months:
    for date,min in ave_dict.items():
      month=date[7:10]
      if month == month_name:
        real = date.replace(month_name,"")
        times.append(f"{real} Total : {round(min/60,1)} hour")
        total+=round(min/60,1)
        x+=1
        y=1
    if x >= 1:
      all_info[month_name] = times
      times = []

    if y==1:
      Total_time[month_name]=round(total,2)
      y=0
    total=0

  return all_info , Total_time

def Top10():
  '''
  return top 10 day , continues
  '''
  Top_in_day=[]
  Top_in_time=[]

  Date_Hour = average()
  set1 = set()
  duplicate = {round(h1/60,2):d for d,h1 in Date_Hour.items()}
  for hour,date in duplicate.items():
    set1.add(hour)
  lis=list(set1)
  lis.sort(reverse=True)
  a=1
  for i in range(10):
    # print(f"{a}) Date : {duplicate[lis[i]]}  Total time : {lis[i]}\n")
    Top_in_day.append(f"{a}) Date : {duplicate[lis[i]]}  Total time : {lis[i]}")
    a+=1
  
  f=open("Timmer.txt")
  Date_Hour={}
  lis1=[]
  s = set()
  for line in f:
    date = line[0:10]
    hour1 = line.split("and")
    final_h = hour1[1].split("H")
    s.add(final_h[0])
    Date_Hour[final_h[0]]=date
  lis1=list(s)
  lis1.sort(reverse=True)
  a=1
  print(lis1)
  for i in range(10):
  #  print(f"{a}) Date : {Date_Hour[lis1[i]]}  Total time : {lis1[i]}\n")
   Top_in_time.append(f"{a}) Date : {Date_Hour[lis1[i]]}  Total time : {lis1[i]} Hour")
   a+=1
  return Top_in_day,Top_in_time

def Topper():
  '''
  return Topper of day and continues time
  '''
  Top_in_day=""
  Top_in_time=""

  Date_Hour = average()
  set1 = set()
  duplicate = {round(h1/60,2):d for d,h1 in Date_Hour.items()}
  for hour,date in duplicate.items():
    set1.add(hour)
  lis=list(set1)
  lis.sort(reverse=True)
  a=1
  Top_in_day=f"Day >> {duplicate[lis[0]][3:]} : {lis[0]} Hour"


  f=open("Timmer.txt")
  Date_Hour={}
  lis1=[]
  for line in f:
    date = line[0:10]
    hour1 = line.split("and")
    final_h = hour1[1].split("H")
    lis1.append(final_h[0])
    Date_Hour[final_h[0]]=date
  lis1.sort(reverse=True)

  Top_in_time=f"Continues >> {Date_Hour[lis1[0]][3:]} : {lis1[0]} Hour"

  return Top_in_day,Top_in_time


#----------------------------------------------------------------------------

def speak(s):
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(s)

# All GUI functions
class GUI(Tk):

    def __init__(self,h,w,title,*argv) -> None:
        '''
        higth , width , Title , min , max
        '''
        super().__init__()
        self.geometry(f"{w}x{h}")
        self.title(title)
        try:
            self.maxsize(argv[0],argv[1])
            self.minsize(argv[2],argv[3])
        except:pass

    def CreateLable(self,*argv):
        '''
        1) text, font, x position ,y position , bg, fg
        2) text,padx,pady, font, x position ,y position ,bg,fg 
        3) text, font 
        '''
        a=0
        try:
            variable = Label(text=argv[0],font=argv[1],bg=argv[4],fg=argv[5])
            variable.place(x=argv[2],y=argv[3])
            a=1
        except:
            pass
        if a == 0:
            try:
                variable=Label(text=argv[0],padx=argv[1],pady=argv[2],font=argv[3])
                variable.place(x=argv[4],y=argv[5])
                a=1
            except:
                pass
        if a == 0:
                variable=Label(text=argv[0],font=argv[1])
                variable.pack()  
        return variable
    
    def B(self,line,Font,func,master=None,x=None,y=None,ima = None,**argv):
        '''
        Create Button
        text, font, command,optional(master , padx , pady , image) , extra **argv
        
        '''
        try:
            try:
                return Button(master,argv,text=line,font=Font,command=func,padx=x,pady=y,image=ima)
            except:
                return Button(self,argv,text=line,font=Font,command=func,padx=x,pady=y,image=ima)
        except:
            try:
                return Button(master,text=line,font=Font,command=func,padx=x,pady=y,image=ima)
            except:
                return Button(self,text=line,font=Font,command=func,padx=x,pady=y,image=ima)
          
    def jpg(self,filename,*argv):
        '''
        file name,resize x, resize y
        '''
        Final_image =  PIL.ImageTk.PhotoImage(PIL.Image.open(filename).resize((argv[0], argv[1]), PIL.Image.ANTIALIAS))

        return Label(image=Final_image)

    def image(self,filename,*argv):
        '''
        file name,resize x, resize y --> ruturn modified image

        '''
        F_image =  PIL.ImageTk.PhotoImage(PIL.Image.open(filename).resize((argv[0], argv[1]), PIL.Image.ANTIALIAS))
        return F_image

    def CreateMenu(self,Font,name,*argv):
        '''
        1.font,(,enu name set),(text1,func1,text2,fun2.. for add seperater , for cut write c),(...)
        2.
        '''
        myMenu = Menu(self)
        m = Menu(myMenu, tearoff=0)
        i = 0
        b = 0
        count=0
        for i in name:
            for i in range(len(argv)):
                if argv[i+2] == "s":
                    m.add_separator()
                    i+=1
                if argv[i+2] == "c" or argv[i+3] == "c":
                    myMenu.add_cascade(label=name[count],menu=m)
                    count+=1
                else:
                    m.add_command(label=argv[i],command=argv[i+1],font=Font)
                    i+=2
                # except:pass
                if count == len(name):
                    break
        # window.config(menu=myMenu)
        return myMenu

def clear(*argv):
    for i in argv:
        i.destroy()



h,m,s,stop=0,0,0,0

def clock(*t):
    global h,s,m,b,l,stop,date

    try:
        if t[0] == "stop":
            speak(f"Sir you work {h} hours and {m} minute")
            if h<2:speak("Sir hour are less than 2 hours time is very limited improve yourself time is very important")
            stop=1
            l.config(text=f"{str(h).zfill(2)}:{str(m).zfill(2)}:{str(s).zfill(2)}")
            l.destroy()
            b.destroy()
            # data.destroy()
            l = Label(text=f"{str(h).zfill(2)}:{str(m).zfill(2)}:{str(s).zfill(2)}",font="lucida 40 bold")
            l.pack(pady=50)
            f=open("Timmer.txt","a") 
            final_hour = (h*60)+(m%60)
            f.write(f"{date} --> {round(float(m+(s/60)+(h*60)),1)} min and {round(float(final_hour/60),2)} Hour\n")

    except Exception as e:
        s+=1
        if s == 60:
            s=0
            m+=1
        elif m==60:
            m=0
            h+=1
        if h == 2 and m==0 and (s==0 or s==1):speak("Sir 2 hours completed")
        if h == 3 and m==0 and (s==0 or s==1):speak("Sir 3 hours completed congrats")
        if h == 4 and m==0 and (s==0 or s==1):speak("Sir 4 hours completed special congrats")
        if h == 5 and m==0 and (s==0 or s==1):speak("Sir 5 hours completed selute you")
        if h == 6 and m==0 and (s==0 or s==1):speak("Oh my god Sir 6 hours completed i dont have word to explain you make history")
        if h >= 7 and m==0 and (s==0 or s==1):speak(f"Oh my god Sir {h} hours completed you are best programmer you are genius in use time perfectly")
        l.config(text=f"{str(h).zfill(2)}:{str(m).zfill(2)}:{str(s).zfill(2)}")
        l.after(1000,clock)

def Print_timer(type1):
    root = Toplevel(window)
    root.resizable(0,0)
    root.wm_iconbitmap("icon.ico")
    root.geometry("1000x600")
    root.config(bg="white")

    all_info,total_by_month=month()
    Day,Time = Top10()
    # clear(bt1,bt2,bt3,greatest)
    root.maxsize(1700,1000)

    if type1 == "history":
        root.geometry("600x700")
        scrollbar = Scrollbar(root)
        scrollbar.pack(side=RIGHT,fill=Y)
        
        data = Listbox(root, yscrollcommand=scrollbar.set,height=50,font="bold 14",width=50)

        f = open("Timmer.txt")
        lines = f.readlines()
        f.close()
        a = 1
        for item in lines:
            if a == 1:
                data.insert(a,"\n\n")
            data.insert(END,f"{a})  {item}")
            data.insert(END,f"\n")
            a+=1
        data.pack()
        scrollbar.config(command=data.yview)
    
    elif type1 == "Top 10":
        can = Canvas(root, width=100, height=300)
        can.pack(fill=X)
        can.create_line(480,0,480,600)
        dayTop = ""
        for item in Day:
            dayTop+=item+"\n\n"
        root.geometry("1000x600")
        day_data = Label(root,text=dayTop,font="lucida 15 bold")
        day_data.place(x=50,y=100)
        title_d = Label(root,text="Top 10 of day",font="lucida 24 bold")
        title_d.place(x=100,y=30)

        timeTop=""
        for item in Time:
            timeTop+=item+"\n\n"
        time_data = Label(root,text=timeTop,font="lucida 15 bold")
        time_data.place(x=550,y=100)
        title_t = Label(root,text="Top 10 of continues time ",font="lucida 24 bold")
        title_t.place(x=580,y=30)
        
    elif type1 == "Month":
        root.geometry("600x700")
        scrollbar = Scrollbar(root)
        scrollbar.pack(side=RIGHT,fill=Y)
        data = Listbox(root, yscrollcommand=scrollbar.set,height=50,font="bold 16",width=50)
        x = 0
        for month_name,times in all_info.items():
            n = 1
            if times==[]:continue
            else:
                if x == 0:
                    title1 = Label(root,text="Total data of per Month\n",font="bold 18")
                    title1.pack()
                data.insert(END,f"\n")
                data.insert(END,f"---------->  {month_name}  <----------  Total : {total_by_month[month_name]} Hours , Average = {round(total_by_month[month_name]/(len(times)),2)}")
                data.insert(END,f"\n")
                x+=1
            
                for count in range(1,32):
                    for time in times:
                        a = str(count)
                        if len(str(count))==1:  
                            count="0"+a 
                        else:   
                            count=str(count)
                        
                        time1 = time.split("Total")
                        date = time1[0]
                        if date.find(count) == 4:
                            data.insert(END,f"{n}) {time}")
                            n+=1
                            break
                        else:continue
                    count=int(count)
                    count+=1


        data.pack()
        scrollbar.config(command=data.yview)

def Timer_data():
    root = Toplevel(window)
    root.geometry("600x200")
    greatest = Label(root,text="=> Select the type <=",font="lucida 25 bold")
    greatest.pack()
    bt1 = Button(root,text="All history",font="lucida 20 bold",command=partial( Print_timer,"history"))
    bt1.pack(side=LEFT,padx=10)
    bt2 = Button(root,text="Top 10",font="lucida 20 bold",command=partial( Print_timer,"Top 10"))
    bt2.pack(side=LEFT,padx=10)
    bt3 = Button(root,text="Every Month",font="lucida 20 bold",command=partial( Print_timer,"Month"))
    bt3.pack(side=LEFT,padx=10)


def close():
    global stop
    if stop == 0:
        if messagebox.askokcancel("Quit","Timer is now running you realy exit program"):
            window.destroy()
    elif stop == 1:
        window.destroy()

if __name__ == '__main__':
    
    window = GUI(400,300,"Clock")
    window.wm_iconbitmap("icon.ico")
    date = datetime.datetime.now().strftime("%a %d %b") 
    window.resizable(0,0)
    # window.config(bg="gray")
    Continue_T,Day_T=Topper()
    l = Label(text="",font="lucida 40 bold")
    l.pack(pady=20)

    var = clock()

    b = Button(text="Stop",command=partial(clock,"stop"),font="lucida 20 bold",bg="red")
    b.pack(pady=5)
    data = Button(text="Data",command=Timer_data,font="lucida 20 bold",bg="red")
    data.pack(pady=5)

    Day = Label(text=Day_T,font="lucida 12 bold",fg="green")
    Day.pack(pady=10)
    Continue = Label(text=Continue_T,font="lucida 12 bold",fg="green")
    Continue.pack(pady=10)

    window.protocol("WM_DELETE_WINDOW",close)
    window.iconify()
    window.update()
    window.mainloop()