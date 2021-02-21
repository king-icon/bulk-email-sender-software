from tkinter import * 
from PIL import ImageTk
from tkinter import messagebox,filedialog
import os
import pandas as pd 
import emailfuntion
import time 

class BULK_EMAIL:
    def __init__(self,root):
        self.root=root
        self.root.title("bulk email application")
        self.root.geometry("1000x550+200+70")
        self.root.resizable(False,False)
        self.root.config(bg="white")

        #=========icons=======#

        self.email_icon=ImageTk.PhotoImage(file="imgs/email.png")
        self.setting_icon=ImageTk.PhotoImage(file="imgs/setting.png")

        #======titls=====#
        title=Label(self.root,text="bulk email send panel",image=self.email_icon,compound=LEFT,font=("goudy old style",48,"bold"),bg="#222A35",fg="white",anchor="w").place(x=0,y=0,relwidth=1)
        des=Label(self.root,text="use excel fie to send bulk emais at oncce in 1 click, ensure the email column name must be email",font=("calibri (body)",14,),bg="#FFD996",fg="#262626").place(x=0,y=90,relwidth=1)
        btn_setting=Button(self.root,image=self.setting_icon,bd=0,activebackground="#222A35",bg="#222A35",cursor="hand2",command=self.seting_window).place(x=900,y=5)

        #=======radiobutton======#
        self.var_click=StringVar()
        single = Radiobutton(self.root,text="single",value="single",variable=self.var_click,activebackground="white",font=("times new roman",30,"bold"),bg="white",fg="#262626",command=self.check_single_or_bulk).place(x=50,y=150)
        bulk = Radiobutton(self.root,text="bulk",value="bulk",variable=self.var_click,activebackground="white",font=("times new roman",30,"bold"),bg="white",fg="#262626",command=self.check_single_or_bulk).place(x=250,y=150)
        self.var_click.set("single")
        
        #++++++++++++#
        to = Label(self.root,text="To (email addres)",font=("times new roman",18),bg="white").place(x=50,y=250)
        subject = Label(self.root,text="subject",font=("times new roman",18),bg="white").place(x=50,y=300)
        msg = Label(self.root,text="message",font=("times new roman",18),bg="white").place(x=50,y=350)

        self.txt_to=Entry(self.root,font=("times new roman",14),bg="lightyellow")
        self.txt_to.place(x=300,y=250,width=350,height=30)

        self.btn_browse=Button(self.root,command=self.browse_file,text="browse",font=("times new roman",15,"bold"),bg="#8FAADC",fg="#262626",activebackground="#8FAADC",activeforeground="#262626",cursor="hand2",state=DISABLED)
        self.btn_browse.place(x=670,y=250,height=30)


        self.txt_subj=Entry(self.root,font=("times new roman",14),bg="lightyellow")
        self.txt_subj.place(x=300,y=300,width=450,height=30)

        self.txt_msg=Text(self.root,font=("times new roman",12),bg="lightyellow")
        self.txt_msg.place(x=300,y=350,width=500,height=120)

        #-------------------------status----------#
        self.lbl_total= Label(self.root,font=("times new roman",18),bg="white")
        self.lbl_total.place(x=50,y=490)

        self.lbl_sent= Label(self.root,font=("times new roman",18),bg="white",fg="green")
        self.lbl_sent.place(x=300,y=490)

        self.lbl_left= Label(self.root,font=("times new roman",18),bg="white",fg="blue")
        self.lbl_left.place(x=400,y=490)

        self.lbl_failed= Label(self.root,font=("times new roman",18),bg="white",fg="orange")
        self.lbl_failed.place(x=500,y=490)
        

        

        btn_clear=Button(self.root,command=self.clear1,text="clear",font=("times new roman",18,"bold"),bg="#262626",fg="white",activebackground="#262626",activeforeground="white",cursor="hand2").place(x=700,y=480,height=30,width=100)
        btn_send=Button(self.root,command=self.send_email_button,text="send msg",font=("times new roman",18,"bold"),bg="#00B0F0",fg="white",activebackground="#00B0F0",activeforeground="white",cursor="hand2").place(x=820,y=480,height=30)

        self.check_file_exists()


    def send_email_button(self):
        x = len(self.txt_msg.get('1.0',END))
        if self.txt_to.get()=="" or self.txt_subj.get()=="" or x==1:
            
            messagebox.showerror("error","all filds are required",parent=self.root)
        else:
            if self.var_click.get()=="single":
                status=emailfuntion.email_send_function(self.txt_to.get(),self.txt_subj.get(),self.txt_msg.get('1.0',END),self.from_,self.pass_)
                if status=="s":
                    messagebox.showinfo("success","emailas been sent",parent=self.root)
                if status=="f":
                    messagebox.showerror("failed","emails not sent,try again ",parent=self.root)


            if self.var_click.get()=="bulk":
                self.failed=[]
                self.s_count=0
                self.f_count=0
                for x in self.emails:
                    status=emailfuntion.email_send_function(x,self.txt_subj.get(),self.txt_msg.get('1.0',END),self.from_,self.pass_)
                    if status=="s":
                        self.s_count+=1
                    if status=="f":
                        self.f_count+=1
                    self.status_bar()
                time.sleep(1)
                messagebox.showinfo("success","emailas been sent, please check status",parent=self.root)

    def status_bar(self):
        self.lbl_total.config(text="status: "+str(len(self.emails))+">>")
        self.lbl_sent.config(text="sent: "+str(self.s_count))
        self.lbl_left.config(text="left: "+str(len(self.emails)-(self.s_count+self.f_count)))
        self.lbl_failed.config(text="failed: "+str(self.f_count))
        self.lbl_total.update()
        self.lbl_sent.update()
        self.lbl_left.update()
        self.lbl_failed.update()


    def check_single_or_bulk(self): #btn check krnek liye single ho to bowse diable ho  jaye
        if self.var_click.get()=="single":
            self.btn_browse.config(state=DISABLED)
            self.txt_to.config(state=NORMAL)                                         #bulk ho to bowse btn chalu ho jaye
            self.txt_to.delete(0,END) 
            self.clear1( )
        if self.var_click.get()=="bulk":
            self.btn_browse.config(state=NORMAL)
            self.txt_to.delete(0,END) 
            self.txt_to.config(state='readonly')                                         #bulk ho to bowse btn chalu ho jaye
            

    def clear1(self): 
        self.txt_to.config(state=NORMAL)                                 #data ko clear krne k liye 
        self.txt_to.delete(0,END)
        self.txt_subj.delete(0,END)
        self.txt_msg.delete('1.0',END)
        self.var_click.set("single")  #clear hone k baad btn sinle ho jaye
        self.btn_browse.config(state=DISABLED)
        self.lbl_total.config(text="")
        self.lbl_sent.config(text="")
        self.lbl_left.config(text="")
        self.lbl_failed.config(text="")                                        #or browse btn diale ho jae 


    def clear2(self):
        self.txt_from.delete(0,END)
        self.txt_pass.delete(0,END)

    def browse_file(self):
        op = filedialog.askopenfile(initialdir='/',title="select excel file for emails",filetypes=(("all files","*"),("Excel fles",".xlsx")))
        if op!=None:
            data = pd.read_excel(op.name)
            
            if "Email" in data.columns:
                self.emails=list(data['Email'])
                c=[]
                for i in self.emails:
                    if pd.isnull(i)==False:
                        c.append(i)
                self.emails=c
                if len(self.emails)>0:
                    self.txt_to.config(state=NORMAL)
                    self.txt_to.delete(0,END)
                    self.txt_to.insert(0,str(op.name.split("/")[-1]))
                    self.txt_to.config(state='readonly')
                    self.lbl_total.config(text="total: "+str(len(self.emails)))
                    self.lbl_sent.config(text="sent: ")
                    self.lbl_left.config(text="left: ")
                    self.lbl_failed.config(text="failed: ")
                    

                else:
                    messagebox.showerror("error","this file does not have email column",parent=self.root)

            else:
                messagebox.showerror("error","please selct a file which contain emails",parent=self.root)

    def seting_window(self):
        self.check_file_exists()
        self.root2=Toplevel() #child class
        self.root2.title("setting")
        self.root2.geometry("700x350+350+120")
        self.root2.config(bg="white")
        self.root2.focus_force()  #2nd screen ko fcus krne k lie
        self.root2.grab_set() #2nd screen ko hide hone se rokne k liye
        title2=Label(self.root2,text="crenditials settings",image=self.setting_icon,compound=LEFT,font=("goudy old style",48,"bold"),bg="#222A35",fg="white",anchor="w").place(x=0,y=0,relwidth=1)
        des2=Label(self.root2,text="enter email adress or password from which to send the all email",font=("calibri (body)",14,),bg="#FFD996",fg="#262626").place(x=0,y=90,relwidth=1)
        
        from_= Label(self.root2,text="enter email",font=("times new roman",18),bg="white").place(x=50,y=150)
        pass_ = Label(self.root2,text="password",font=("times new roman",18),bg="white").place(x=50,y=200)

        self.txt_from=Entry(self.root2,font=("times new roman",14),bg="lightyellow")
        self.txt_from.place(x=250,y=150,width=350,height=30)

        self.txt_pass=Entry(self.root2,font=("times new roman",14),bg="lightyellow",show="*")
        self.txt_pass.place(x=250,y=200,width=350,height=30)

        btn_clear2=Button(self.root2,command=self.clear2,text="clear",font=("times new roman",18,"bold"),bg="#262626",fg="white",activebackground="#262626",activeforeground="white",cursor="hand2").place(x=300,y=260,height=30,width=100)
        btn_save=Button(self.root2,command=self.save_setting,text="save",font=("times new roman",18,"bold"),bg="#00B0F0",fg="white",activebackground="#00B0F0",activeforeground="white",cursor="hand2").place(x=430,y=260,height=30)
        
        self.txt_from.insert(0,self.from_)
        self.txt_pass.insert(0,self.pass_)

    def save_setting(self):
        if self.txt_from.get()=="" or self.txt_pass.get()=="":
            messagebox.showerror("error","all fields are required",parent=self.root2)
        else:
            f=open('settings.txt','w')
            f.write(self.txt_from.get()+","+self.txt_pass.get())
            f.close()
            messagebox.showinfo("sucess"," data saved sucessfully",parent=self.root2)
            self.check_file_exists()
    def check_file_exists(self):
        if os.path.exists("settings.txt")==False:
            f=open('settings.txt','w')
            f.write(",")
            f.close()
        f2=open('settings.txt','r')
        self.crenditial=[]
        for i in f2:
            self.crenditial.append( [i.split(",")[0],i.split(",")[1]] )
        self.from_=self.crenditial[0][0]
        self.pass_=self.crenditial[0][1]


root = Tk()
obj = BULK_EMAIL(root)
root.mainloop()
