from tkinter import *
from tkinter import ttk
import pandas as pd
def Brand_Name():
    data=pd.read_excel("MOTOR_DATA.xlsx",index_col=0)
    return data["BRAND"].value_counts()
################# Brand Information #################
class Brand_Information(Frame):
    def __init__(self,master=None):
        super().__init__(master)
        self.master.configure(background='#f2f6f4')
        self.master.title("DANH SÁCH CÁCH HÃNG XE")
        self.master.geometry("500x350")
        self.pack()
        self.List_Brand()
    def List_Brand(self):
        self.ListBrand=LabelFrame(self.master,text="Danh sách tất cả hãng xe",padx=20, pady=20)
        self.ListBrand.pack(side='top')
        self.ListView=ttk.Treeview(self.ListBrand,height=10)
        self.ListView['columns']=('INDEX','MODEL', "COUNT")

        self.ListView.column("#0",width=0,minwidth=0)
        self.ListView.column("INDEX",anchor=CENTER,width=80)
        self.ListView.column("MODEL",anchor=CENTER,width=180)
        self.ListView.column("COUNT",anchor=CENTER,width=180)

        self.ListView.heading("INDEX",text="STT",anchor=CENTER)
        self.ListView.heading("MODEL",text="HÃNG",anchor=CENTER)
        self.ListView.heading("COUNT",text="SỐ LƯỢNG MẪU MÃ",anchor=CENTER)
        for i in range(1,len(Brand_Name().keys())+1):
            self.ListView.insert(parent='', index='end',values=(i,str(Brand_Name().keys()[i-1]),str(Brand_Name()[i-1])))
        self.ListView.pack(side=TOP)
        self.Quit=Button(self.ListBrand, text="QUIT", fg="red",command=self.master.destroy,width=10,height=2)
        self.Quit.pack(side=BOTTOM)


def List_Brand_func():
    root=Tk()
    Brand_Information_Application=Brand_Information(master=root)
    Brand_Information_Application.mainloop()

################# MODEL Information #################
class Model_Information(Frame):
    def __init__(self,master=None):
        super().__init__(master)
        self.master.configure(background='#f2f6f4')
        self.master.title("DANH SÁCH CÁCH HÃNG XE")
        self.master.geometry("1000x550")
        self.pack()
        self.List_Model()
    def List_Model(self):
        self.ListModel=LabelFrame(self.master,text="Danh sách tất cả hãng xe",padx=20, pady=20)
        self.ListModel.pack(side='top')
        self.ListView=ttk.Treeview(self.ListModel,height=13)
        self.ListView['columns']=('INDEX','BRAND', "MODEL",'ENGINE','CAPACITY', "FUEL",'FUEL TANK CAPACITY','PRICE')
    
        self.ListView.column("#0",width=0,minwidth=0)
        self.ListView.column("INDEX",anchor=CENTER,width=50)
        self.ListView.column("BRAND",anchor=CENTER,width=80)
        self.ListView.column("MODEL",anchor=CENTER,width=120)
        self.ListView.column("ENGINE",anchor=CENTER,width=100)
        self.ListView.column("CAPACITY",anchor=CENTER,width=100)
        self.ListView.column("FUEL",anchor=CENTER,width=180)
        self.ListView.column("FUEL TANK CAPACITY",anchor=CENTER,width=180)
        self.ListView.column("PRICE",anchor=CENTER,width=100)

        self.ListView.heading("INDEX",text="STT",anchor=CENTER)
        self.ListView.heading("BRAND",text="HÃNG",anchor=CENTER)
        self.ListView.heading("MODEL",text="MẪU",anchor=CENTER)
        self.ListView.heading("ENGINE",text="ĐỘNG CƠ",anchor=CENTER)
        self.ListView.heading("CAPACITY",text="DUNG TÍCH ĐỘNG CƠ",anchor=CENTER)
        self.ListView.heading("FUEL",text="TIÊU THỤ",anchor=CENTER)
        self.ListView.heading("FUEL TANK CAPACITY",text="DUNG TÍCH NHIÊN LIỆU",anchor=CENTER)
        self.ListView.heading("PRICE",text="GIÁ",anchor=CENTER)
        data=pd.read_excel("MOTOR_DATA.xlsx",index_col=0)
        for i in range(1,len(data)+1):
            self.ListView.insert(parent='', index='end',values=(i,str(data.iloc[i-1][0]),str(data.iloc[i-1][1]),str(data.iloc[i-1][2]),str(data.iloc[i-1][4]),str(data.iloc[i-1][6]),str(data.iloc[i-1][7]),str(data.iloc[i-1][9])))
        self.ListView.pack(side=TOP)
        self.Quit=Button(self.ListModel, text="QUIT", fg="red",command=self.master.destroy,width=10,height=2,)
        self.Quit.pack(side=BOTTOM)
        self.Sum=Label(self.ListModel, text="TỔNG CỘNG: "+str(len(data)),padx=5,pady=5)
        self.Sum.pack(side=LEFT)



def List_Model_func():
    root=Tk()
    Model_Information_Application=Model_Information(master=root)
    Model_Information_Application.mainloop()


################# ADD MODEL AND BRAND #################
def Check_and_Add_Model(Price,Amount,TankCapacity,Fuel,Brand,Model,Engine,Technology,Capacity,LWH):
        data=pd.read_excel("MOTOR_DATA.xlsx",index_col=0)
        try:
            data.loc[len(data)+1]=[str(Brand.get()),str(Model.get()),str(Engine.get()),
            str(Technology.get()),str(Capacity.get()),str(LWH.get()),str(Fuel.get()),
            str(TankCapacity.get()), int(Amount.get()),int(Price.get())]
            data.to_excel("MOTOR_DATA.xlsx")
        except:
            root=Tk()
            L1 = Label(root, text="Thất bại do lỗi nhập")
            L1.pack()
            Quit=Button(root, text="QUIT", fg="red",command=root.destroy,width=10,height=2)
            Quit.pack()
            L1.mainloop()
        else:
            root=Tk()
            L1 = Label(root, text="Đã Lưu")
            L1.pack()
            Quit=Button(root, text="QUIT", fg="red",command=root.destroy,width=10,height=2)
            Quit.pack()
            L1.mainloop()
class Add_Model(Frame):
    def __init__(self,master=None):
        super().__init__(master)
        self.master.configure(background='#f2f6f4')
        self.master.title("HỆ THỐNG QUẢN LÝ XE MÁY")
        self.master.geometry("500x500")
        self.pack()
        self.add_new_model()
    def add_new_model(self):
        self.Add_New_Model=LabelFrame(self.master,text="THÊM SẢN PHẨM VÀ HÃNG XE",padx=20, pady=20)
        self.Add_New_Model.pack(side='top')
        self.Brand=Label(self.Add_New_Model,text="Hãng Xe:",padx=5,pady=5)
        self.Brand.grid(column=0,row=1)
        self.Model=Label(self.Add_New_Model,text="Loại Xe:",padx=5,pady=5)
        self.Model.grid(column=0,row=2)
        self.Engine=Label(self.Add_New_Model,text="Động cơ:",padx=5,pady=5)
        self.Engine.grid(column=0,row=3)
        self.Technology=Label(self.Add_New_Model,text="Công Nghệ:",padx=5,pady=5)
        self.Technology.grid(column=0,row=4)
        self.Capacity=Label(self.Add_New_Model,text="Dung tích Xilanh:",padx=5,pady=5)
        self.Capacity.grid(column=0,row=5)
        self.LWH=Label(self.Add_New_Model,text="Kích Thước:",padx=5,pady=5)
        self.LWH.grid(column=0,row=6)
        self.Fuel=Label(self.Add_New_Model,text="Tiêu thụ nhiên Liệu/100km:",padx=5,pady=5)
        self.Fuel.grid(column=0,row=7)
        self.TankCapacity=Label(self.Add_New_Model,text="Dung tích nhiên liệu:",padx=5,pady=5)
        self.TankCapacity.grid(column=0,row=8)
        self.Amount=Label(self.Add_New_Model,text="Số lượng nhập về:",padx=5,pady=5)
        self.Amount.grid(column=0,row=9)
        self.Price=Label(self.Add_New_Model,text="Giá:",padx=5,pady=5)
        self.Price.grid(column=0,row=10)

        Brand=Entry(self.Add_New_Model)
        Brand.grid(column=1,row=1)
        Model=Entry(self.Add_New_Model)
        Model.grid(column=1,row=2)
        Engine=Entry(self.Add_New_Model)
        Engine.grid(column=1,row=3)
        Technology=Entry(self.Add_New_Model)
        Technology.grid(column=1,row=4)
        Capacity=Entry(self.Add_New_Model)
        Capacity.grid(column=1,row=5)
        LWH=Entry(self.Add_New_Model)
        LWH.grid(column=1,row=6)
        Fuel=Entry(self.Add_New_Model)
        Fuel.grid(column=1,row=7)
        TankCapacity=Entry(self.Add_New_Model)
        TankCapacity.grid(column=1,row=8)
        Amount=Entry(self.Add_New_Model)
        Amount.grid(column=1,row=9)
        Price=Entry(self.Add_New_Model)
        Price.grid(column=1,row=10)

        self.Quit=Button(self.Add_New_Model, text="QUIT", fg="red",command=self.master.destroy,width=10,height=2)
        self.Quit.grid(column=1,row=12)

        self.Save=Button(self.Add_New_Model, text="SAVE", fg="green",command= lambda: Check_and_Add_Model(Price,Amount,TankCapacity,Fuel,Brand,Model,Engine,Technology,Capacity,LWH),width=10,height=2)
        self.Save.grid(column=0,row=12)




def Add_Model_func():
    root=Tk()
    Add_Model_Application=Add_Model(master=root)
    Add_Model_Application.mainloop()

################# DELETE MODEL #################
def Check_and_Delete_MODEL(variable):
    try:
        data= pd.read_excel("MOTOR_DATA.xlsx",index_col=0)
        Index=list(data.loc[data['TYPE'] == variable].index)
        data=data.drop(Index)
    except:
        data.to_excel("MOTOR_DATA.xlsx")
        root=Tk()
        root.geometry("100x100")
        L1 = Label(root, text="Lỗi")
        L1.pack()
        Quit=Button(root, text="QUIT", fg="red",command=root.destroy,width=10,height=2)
        Quit.pack()
        L1.mainloop()
    else:
        data.to_excel("MOTOR_DATA.xlsx")
        root=Tk()
        root.geometry("100x100")
        L1 = Label(root, text="Đã Lưu")
        L1.pack()
        Quit=Button(root, text="QUIT", fg="red",command=root.destroy,width=10,height=2)
        Quit.pack()
        L1.mainloop()
    
class Del_Model(Frame):
    def __init__(self,master=None):
        super().__init__(master)
        self.master.configure(background='#f2f6f4')
        self.master.title("HỆ THỐNG QUẢN LÝ XE MÁY")
        self.master.geometry("500x500")
        self.pack()
        self.del_model()
    def del_model(self):
        self.Del_Model=LabelFrame(self.master,text="XÓA MẪU XE",padx=20, pady=20)
        self.Del_Model.pack(side='top')
        data= pd.read_excel("MOTOR_DATA.xlsx",index_col=0)
        OPTIONS= [data.iloc[i][1] for i in range(0,len(data))]
        variable = StringVar(self.Del_Model)
        variable.set(OPTIONS[0]) # default value
        w = OptionMenu(self.Del_Model, variable, *OPTIONS)
        w.grid(column=1,row=0)
        self.Quit=Button(self.Del_Model, text="QUIT", fg="red",command=self.master.destroy,width=10,height=2)
        self.Quit.grid(column=2,row=1)
        self.Save=Button(self.Del_Model, text="SAVE", fg="green",command= lambda: Check_and_Delete_MODEL(variable.get()),width=10,height=2)
        self.Save.grid(column=0,row=1)



def Del_Model_func():
    root=Tk()
    Application=Del_Model(master=root)
    Application.mainloop()

    ################# DELETE BRAND #################
def Check_and_Delete_BRAND(variable):
    try:
        data= pd.read_excel("MOTOR_DATA.xlsx",index_col=0)
        Index=list(data.loc[data['BRAND'] == variable].index)
        data=data.drop(Index)
        data.to_excel("MOTOR_DATA.xlsx")
    except:
        root=Tk()
        root.geometry("100x100")
        L1 = Label(root, text="Lỗi")
        L1.pack()
        Quit=Button(root, text="QUIT", fg="red",command=root.destroy,width=10,height=2)
        Quit.pack()
        L1.mainloop()
    else:
        root=Tk()
        root.geometry("100x100")
        L1 = Label(root, text="Đã Lưu")
        L1.pack()
        Quit=Button(root, text="QUIT", fg="red",command=root.destroy,width=10,height=2)
        Quit.pack()
        L1.mainloop()
    
class Del_Brand(Frame):
    def __init__(self,master=None):
        super().__init__(master)
        self.master.configure(background='#f2f6f4')
        self.master.title("HỆ THỐNG QUẢN LÝ XE MÁY")
        self.master.geometry("500x500")
        self.pack()
        self.del_Brand()
    def del_Brand(self):
        self.Del_Brand=LabelFrame(self.master,text="XÓA NHÃN HÀNG",padx=20, pady=20)
        self.Del_Brand.pack(side='top')
        data= pd.read_excel("MOTOR_DATA.xlsx",index_col=0)
        OPTIONS= list(data["BRAND"].value_counts().keys())
        variable = StringVar(self.Del_Brand)
        variable.set(OPTIONS[0]) # default value
        w = OptionMenu(self.Del_Brand, variable, *OPTIONS)
        w.grid(column=1,row=0)
        self.Quit=Button(self.Del_Brand, text="QUIT", fg="red",command=self.master.destroy,width=10,height=2)
        self.Quit.grid(column=2,row=1)
        self.Save=Button(self.Del_Brand, text="SAVE", fg="green",command= lambda: Check_and_Delete_BRAND(variable.get()),width=10,height=2)
        self.Save.grid(column=0,row=1)



def Del_Brand_func():
    root=Tk()
    Application=Del_Brand(master=root)
    Application.mainloop()

    ################# EDIT BRAND #################
def Check_and_edit_Brand(variable,Brand):
    try:
        data= pd.read_excel("MOTOR_DATA.xlsx",index_col=0)
        Index=list(data.loc[data['BRAND'] == variable].index)
        for i in Index:
            data.loc[i,'BRAND']=Brand
        data.to_excel("MOTOR_DATA.xlsx")
    except:
        root=Tk()
        root.geometry("100x100")
        L1 = Label(root, text="Lỗi")
        L1.pack()
        Quit=Button(root, text="QUIT", fg="red",command=root.destroy,width=10,height=2)
        Quit.pack()
        L1.mainloop()
    else:
        data.to_excel("MOTOR_DATA.xlsx")
        root=Tk()
        root.geometry("100x100")
        L1 = Label(root, text="Đã Lưu")
        L1.pack()
        Quit=Button(root, text="QUIT", fg="red",command=root.destroy,width=10,height=2)
        Quit.pack()
        L1.mainloop()

class Edit_Brand(Frame):
    def __init__(self,master=None):
        super().__init__(master)
        self.master.configure(background='#f2f6f4')
        self.master.title("HỆ THỐNG QUẢN LÝ XE MÁY")
        self.master.geometry("500x500")
        self.pack()
        self.Edit_Brand()
    def Edit_Brand(self):
        self.Edit_Brand=LabelFrame(self.master,text="SỬA NHÃN HÀNG",width=300,height=300,padx=20, pady=20)
        self.Edit_Brand.pack(side='top')
        data1= pd.read_excel("MOTOR_DATA.xlsx",index_col=0)
        OPTIONS= list(data1["BRAND"].value_counts().keys())
        variable = StringVar(self.Edit_Brand)
        variable.set(OPTIONS[0]) # default value
        w = OptionMenu(self.Edit_Brand, variable, *OPTIONS)
        w.grid(column=0,row=0)

        self.Brand=Label(self.Edit_Brand,text="Đổi thành:",padx=5,pady=5)
        self.Brand.grid(column=0,row=1)
        Brand=Entry(self.Edit_Brand)
        Brand.grid(column=1,row=1)

        self.Quit=Button(self.Edit_Brand, text="QUIT", fg="red",command=self.master.destroy,width=10,height=2)
        self.Quit.grid(column=1,row=3)
        self.Save=Button(self.Edit_Brand, text="SAVE", fg="green",command= lambda: Check_and_edit_Brand(variable.get(),Brand.get()),width=10,height=2)
        self.Save.grid(column=0,row=3)



def Edit_Brand_func():
    root=Tk()
    Application=Edit_Brand(master=root)
    Application.mainloop()


    ################# EDIT Model #################
def Check_and_edit(variable,Model):
    try:
        data= pd.read_excel("MOTOR_DATA.xlsx",index_col=0)
        Index=list(data.loc[data['TYPE'] == variable].index)
        for i in Index:
            data.loc[i,'TYPE']=Model
        data.to_excel("MOTOR_DATA.xlsx")
    except:
        root=Tk()
        root.geometry("100x100")
        L1 = Label(root, text="Lỗi")
        L1.pack()
        Quit=Button(root, text="QUIT", fg="red",command=root.destroy,width=10,height=2)
        Quit.pack()
        L1.mainloop()
    else:
        root=Tk()
        root.geometry("100x100")
        L1 = Label(root, text="Đã Lưu")
        L1.pack()
        Quit=Button(root, text="QUIT", fg="red",command=root.destroy,width=10,height=2)
        Quit.pack()
        L1.mainloop()

class Edit_Model(Frame):
    def __init__(self,master=None):
        super().__init__(master)
        self.master.configure(background='#f2f6f4')
        self.master.title("HỆ THỐNG QUẢN LÝ XE MÁY")
        self.master.geometry("500x500")
        self.pack()
        self.Edit_Model()
    def Edit_Model(self):
        self.Edit_Model=LabelFrame(self.master,text="SỬA TÊN MODEL",width=300,height=300,padx=20, pady=20)
        self.Edit_Model.pack(side='top')
        data= pd.read_excel("MOTOR_DATA.xlsx",index_col=0)
        OPTIONS= [data.iloc[i][1] for i in range(0,len(data))]
        variable = StringVar(self.Edit_Model)
        variable.set(OPTIONS[0]) # default value
        w = OptionMenu(self.Edit_Model, variable, *OPTIONS)
        w.grid(column=0,row=0)

        self.Model=Label(self.Edit_Model,text="Đổi thành:",padx=5,pady=5)
        self.Model.grid(column=0,row=1)
        Model=Entry(self.Edit_Model)
        Model.grid(column=1,row=1)

        self.Quit=Button(self.Edit_Model, text="QUIT", fg="red",command=self.master.destroy,width=10,height=2)
        self.Quit.grid(column=1,row=3)
        self.Save=Button(self.Edit_Model, text="SAVE", fg="green",command= lambda: Check_and_edit(variable.get(),Model.get()),width=10,height=2)
        self.Save.grid(column=0,row=3)

def Edit_Model_func():
    root=Tk()
    Application=Edit_Model(master=root)
    Application.mainloop()

    ################# ORDERS INFORMATION #################
def Inputdata(self,Month):
        data=pd.read_excel("ORDERS.xlsx",index_col=0)
        for i in range(1,len(data)+1):
            if(str(data.iloc[i-1][2])==str(Month)):
                self.ListView.insert(parent='', index='end',values=(str(data.iloc[i-1][0]),
                str(data.iloc[i-1][1]),str(data.iloc[i-1][2]),str(data.iloc[i-1][3]),str(data.iloc[i-1][4]),
                str(data.iloc[i-1][5]),str(data.iloc[i-1][6]),str(round(data.iloc[i-1][7])),str(round(data.iloc[i-1][8]))))
        self.ListView.pack(side=TOP)
        self.Quit=Button(self.ListOrders, text="QUIT", fg="red",command=self.master.destroy,width=10,height=2,)
        self.Quit.pack(side=BOTTOM)
        self.Sum=Label(self.ListOrders, text="TỔNG CỘNG: "+str(len(data)),padx=5,pady=5)
        self.Sum.pack(side=LEFT)

class Orders_Information(Frame):
    def __init__(self,master=None):
        super().__init__(master)
        self.master.configure(background='#f2f6f4')
        self.master.title("HỆ THỐNG QUẢN LÝ XE MÁY")
        self.master.geometry("1000x550")
        self.pack()
        self.List_Orders()
    def List_Orders(self):
        self.ListOrders=LabelFrame(self.master,text="Danh sách đơn hàng ",padx=20, pady=20)
        self.ListOrders.pack(side='top')
        self.ListView=ttk.Treeview(self.ListOrders,height=18)
        self.ListView['columns']=('ID','DAY', "MONTH",'YEAR','BRAND', "MODEL",'AMOUNT','PRICE','TOTAL')
    
        self.ListView.column("#0",width=0,minwidth=0)
        self.ListView.column("ID",anchor=CENTER,width=50)
        self.ListView.column("DAY",anchor=CENTER,width=50)
        self.ListView.column("MONTH",anchor=CENTER,width=50)
        self.ListView.column("YEAR",anchor=CENTER,width=100)
        self.ListView.column("BRAND",anchor=CENTER,width=180)
        self.ListView.column("MODEL",anchor=CENTER,width=180)
        self.ListView.column("AMOUNT",anchor=CENTER,width=100)
        self.ListView.column("PRICE",anchor=CENTER,width=100)
        self.ListView.column("TOTAL",anchor=CENTER,width=100)

        self.ListView.heading("ID",text="MÃ ID",anchor=CENTER)
        self.ListView.heading("DAY",text="NGÀY",anchor=CENTER)
        self.ListView.heading("MONTH",text="THÁNG",anchor=CENTER)
        self.ListView.heading("YEAR",text="NĂM",anchor=CENTER)
        self.ListView.heading("BRAND",text="NHÃN HÀNG",anchor=CENTER)
        self.ListView.heading("MODEL",text="MODEL",anchor=CENTER)
        self.ListView.heading("AMOUNT",text="SỐ LƯỢNG",anchor=CENTER)
        self.ListView.heading("PRICE",text="GIÁ",anchor=CENTER)
        self.ListView.heading("TOTAL",text="TỔNG CỘNG",anchor=CENTER)


        
def Orders_Info_func(Month):
    root=Tk()
    Application=Orders_Information(root)
    Inputdata(Application, Month)
    Application.mainloop()

    ################# FIND ORDERS INFORMATION #################
def Show_Information(ID):
    
    data=pd.read_excel("ORDERS.xlsx",index_col=0)
    Index=list(data.loc[data['ID'] == int(ID)].index)
    root=Tk()
    root.geometry("400x200")
    root.title("HỆ THỐNG QUẢN LÝ XE MÁY")
    for i in Index:
        L1 = Label(root, text="Nhãn hàng: "+ str(data.iloc[i-1][4]),padx=5,pady=5)
        L1.pack()
        L2 = Label(root, text="Model: "+ str(data.iloc[i-1][5]),padx=5,pady=5)
        L2.pack()
        L3 = Label(root, text="Số lượng: "+ str(data.iloc[i-1][6]),padx=5,pady=5)
        L3.pack()
        L4 = Label(root, text="Giá: "+ str(round(data.iloc[i-1][7])),padx=5,pady=5)
        L4.pack()
        L5 = Label(root, text="Tổng cộng: "+ str(round(data.iloc[i-1][8])),padx=5,pady=5)
        L5.pack()
    Quit=Button(root, text="QUIT", fg="red",command=root.destroy,width=10,height=2,padx=10,pady=10)
    Quit.pack()
    root.mainloop()

    
class Find_Orders_Information(Frame):
    def __init__(self,master=None):
        super().__init__(master)
        self.master.configure(background='#f2f6f4')
        self.master.title("HỆ THỐNG QUẢN LÝ XE MÁY")
        self.master.geometry("400x200")
        self.pack()
        self.Find_Orders_Info()
    def Find_Orders_Info(self):
        self.Find_Orders=LabelFrame(self.master,text="Tìm Thông tin Đơn Hàng",width=300,height=300,padx=20, pady=20)
        self.Find_Orders.pack(side='top')
        self.ID=Label(self.Find_Orders,text="Mã ID Đơn Hàng:",padx=5,pady=5)
        self.ID.grid(column=0,row=1)
        ID=Entry(self.Find_Orders)
        ID.grid(column=1,row=1)

        self.Quit=Button(self.Find_Orders, text="QUIT", fg="red",command = self.master.destroy,width=10,height=2)
        self.Quit.grid(column=1,row=3)
        self.Save=Button(self.Find_Orders, text="FIND", fg="green",command= lambda: Show_Information(ID.get()),width=10,height=2)
        self.Save.grid(column=0,row=3)

def Find_Orders_Info_func():
    root=Tk()
    Application=Find_Orders_Information(master=root)
    Application.mainloop()


    ################# EDIT ORDERS INFORMATION #################
def Edit_and_Check_Orders_Info(ID,Brand,Model,Amount,Price):
    try:
        data_motor= pd.read_excel("MOTOR_DATA.xlsx",index_col=0)
        data_orders= pd.read_excel("ORDERS.xlsx",index_col=0)
        Index=list(data_orders.loc[data_orders['ID'] == int(ID)].index)
        data_orders.loc[Index[0],'BRAND']= Brand
        data_orders.loc[Index[0],'MODEL']= Model
        data_orders.loc[Index[0],'AMOUNT']= int(Amount)
        data_orders.loc[Index[0],'PRICE']= int(Price)
        data_orders.loc[Index[0],'TOTAL']= int(Price)*int(Amount)
        data_orders.to_excel("ORDERS.xlsx")
    except:
        root=Tk()
        root.geometry("100x100")
        L1 = Label(root, text="Lỗi")
        L1.pack()
        Quit=Button(root, text="QUIT", fg="red",command=root.destroy,width=10,height=2)
        Quit.pack()
        L1.mainloop()
    else:
        root=Tk()
        root.geometry("100x100")
        L1 = Label(root, text="Đã Lưu")
        L1.pack()
        Quit=Button(root, text="QUIT", fg="red",command=root.destroy,width=10,height=2)
        Quit.pack()
        L1.mainloop()


def Show_and_Edit_Information(ID):
    
    data=pd.read_excel("ORDERS.xlsx",index_col=0)
    Index=list(data.loc[data['ID'] == int(ID)].index)
    
    root=Tk()
    root.geometry("300x250")
    root.title("HỆ THỐNG QUẢN LÝ XE MÁY")
    for i in Index:
        L1 = Label(root, text="Nhãn hàng: "+ str(data.iloc[i-1][4]),padx=5,pady=5)
        L1.grid(column=0,row=0)
        E1=Entry(root)
        L2 = Label(root, text="Model: "+ str(data.iloc[i-1][5]),padx=5,pady=5)
        L2.grid(column=0,row=1)
        L3 = Label(root, text="Số lượng: "+ str(data.iloc[i-1][6]),padx=5,pady=5)
        L3.grid(column=0,row=2)
        E3=Entry(root)
        E3.grid(column=1,row=2)
        L4 = Label(root, text="Giá: "+ str(round(data.iloc[i-1][7])),padx=5,pady=5)
        L4.grid(column=0,row=3)
        E4=Entry(root)
        E4.grid(column=1,row=3)

        data= pd.read_excel("MOTOR_DATA.xlsx",index_col=0)
        OPTIONS= list(data["BRAND"].value_counts().keys())
        variable1 = StringVar(root)
        variable1.set(OPTIONS[0]) # default value
        w = OptionMenu(root, variable1, *OPTIONS)
        w.grid(column=1,row=0)


        OPTIONS= [data.iloc[i][1] for i in range(0,len(data))]
        variable2 = StringVar(root)
        variable2.set(OPTIONS[0]) # default value
        w = OptionMenu(root, variable2, *OPTIONS)
        w.grid(column=1,row=1)
    Quit=Button(root, text="QUIT", fg="red",command=root.destroy,width=10,height=2,padx=10,pady=10)
    Quit.grid(column=1,row=5)
    Save=Button(root, text="SAVE", fg="green",command=lambda:Edit_and_Check_Orders_Info(ID,variable1.get(),variable2.get(),E3.get(),E4.get()) ,width=10,height=2,padx=10,pady=10)
    Save.grid(column=0,row=5)
    root.mainloop()

    
class Edit_Orders_Information(Frame):
    def __init__(self,master=None):
        super().__init__(master)
        self.master.configure(background='#f2f6f4')
        self.master.title("HỆ THỐNG QUẢN LÝ XE MÁY")
        self.master.geometry("400x200")
        self.pack()
        self.Edit_Orders_Info()
    def Edit_Orders_Info(self):
        self.Edit_Orders=LabelFrame(self.master,text="Tìm Thông tin Đơn Hàng",width=300,height=300,padx=20, pady=20)
        self.Edit_Orders.pack(side='top')
        self.ID=Label(self.Edit_Orders,text="Mã ID Đơn Hàng:",padx=5,pady=5)
        self.ID.grid(column=0,row=1)
        ID=Entry(self.Edit_Orders)
        ID.grid(column=1,row=1)



        self.Quit=Button(self.Edit_Orders, text="QUIT", fg="red",command = self.master.destroy,width=10,height=2)
        self.Quit.grid(column=1,row=3)
        self.Save=Button(self.Edit_Orders, text="Edit", fg="green",command= lambda: Show_and_Edit_Information(ID.get()),width=10,height=2)
        self.Save.grid(column=0,row=3)

def Edit_Orders_Info_func():
    root=Tk()
    Application=Edit_Orders_Information(master=root)
    Application.mainloop()
###################### Delete Orders ##################


def Check_and_Delete_Orders(ID):
    try:
        data= pd.read_excel("ORDERS.xlsx",index_col=0)
        Index=list(data.loc[data['ID'] == int(ID)].index)
        data=data.drop(Index)
        data.to_excel("ORDERS.xlsx")
    except:
        root=Tk()
        root.geometry("100x100")
        L1 = Label(root, text="Lỗi")
        L1.pack()
        Quit=Button(root, text="QUIT", fg="red",command=root.destroy,width=10,height=2)
        Quit.pack()
        L1.mainloop()
    else:
        root=Tk()
        root.geometry("100x100")
        L1 = Label(root, text="Đã Lưu")
        L1.pack()
        Quit=Button(root, text="QUIT", fg="red",command=root.destroy,width=10,height=2)
        Quit.pack()
        L1.mainloop()
    
class Del_Orders(Frame):
    def __init__(self,master=None):
        super().__init__(master)
        self.master.configure(background='#f2f6f4')
        self.master.title("HỆ THỐNG QUẢN LÝ XE MÁY")
        self.master.geometry("300x150")
        self.pack()
        self.del_Orders()
    def del_Orders(self):
        self.Del_Orders=LabelFrame(self.master,text="XÓA ĐƠN HÀNG",padx=20, pady=20)
        self.Del_Orders.pack(side='top')
        L1=Label(self.Del_Orders,text="Nhập ID Đơn Cần Xóa: ")
        L1.grid(column=0,row=0)
        E1=Entry(self.Del_Orders)
        E1.grid(column=1,row=0)

        self.Quit=Button(self.Del_Orders, text="QUIT", fg="red",command = self.master.destroy,width=10,height=2)
        self.Quit.grid(column=1,row=3)
        self.Save=Button(self.Del_Orders, text="Edit", fg="green",command= lambda: Check_and_Delete_Orders(E1.get()),width=10,height=2)
        self.Save.grid(column=0,row=3)



def Del_Orders_func():
    root=Tk()
    Application=Del_Orders(master=root)
    Application.mainloop()


################### ADD ORDERS ######################

def Check_and_Add_Orders(DAY,MONTH,YEAR,MODEL,AMOUNT):
        data=pd.read_excel("ORDERS.xlsx",index_col=0)
        data_motor= pd.read_excel("MOTOR_DATA.xlsx",index_col=0)
        try:
            data.loc[int(data.iloc[-1][0])+1]=[int(data.iloc[-1][0])+1,
            int(DAY),
            int(MONTH),
            int(YEAR),
            data_motor['BRAND'].values[list(data_motor.loc[data_motor['TYPE'] == str(MODEL)].index)[0]],
            str(MODEL),
            int(AMOUNT),
            int(data_motor['PRICE'].values[list(data_motor.loc[data_motor['TYPE'] == str(MODEL)].index)[0]])*1.15,
            int(int(data_motor['PRICE'].values[list(data_motor.loc[data_motor['TYPE'] == str(MODEL)].index)[0]])*1.15)*int(AMOUNT)
            ]
        except:
            root=Tk()
            L1 = Label(root, text="Thất bại do lỗi nhập hoặc hết hàng")
            L1.pack()
            Quit=Button(root, text="QUIT", fg="red",command=root.destroy,width=10,height=2)
            Quit.pack()
            L1.mainloop()
        else:
            data.to_excel("ORDERS.xlsx")
            root=Tk()
            L1 = Label(root, text="Đã Lưu")
            L1.pack()
            Quit=Button(root, text="QUIT", fg="red",command=root.destroy,width=10,height=2)
            Quit.pack()
            L1.mainloop()

class Add_Orders(Frame):
    def __init__(self,master=None):
        super().__init__(master)
        self.master.configure(background='#f2f6f4')
        self.master.title("HỆ THỐNG QUẢN LÝ XE MÁY")
        self.master.geometry("500x500")
        self.pack()
        self.add_new_Orders()
    def add_new_Orders(self):
        self.Add_New_Orders=LabelFrame(self.master,text="THÊM ĐƠN HÀNG",padx=20, pady=20)
        self.Add_New_Orders.pack(side='top')
        self.DAY=Label(self.Add_New_Orders,text="NGÀY:",padx=5,pady=5)
        self.DAY.grid(column=0,row=1)
        self.MONTH=Label(self.Add_New_Orders,text="THÁNG:",padx=5,pady=5)
        self.MONTH.grid(column=0,row=2)
        self.YEAR=Label(self.Add_New_Orders,text="NĂM:",padx=5,pady=5)
        self.YEAR.grid(column=0,row=3)
        self.MODEL=Label(self.Add_New_Orders,text="MODEL:",padx=5,pady=5)
        self.MODEL.grid(column=0,row=4)
        self.AMOUNT=Label(self.Add_New_Orders,text="SỐ LƯỢNG:",padx=5,pady=5)
        self.AMOUNT.grid(column=0,row=5)

        DAY=Entry(self.Add_New_Orders)
        DAY.grid(column=1,row=1)
        MONTH=Entry(self.Add_New_Orders)
        MONTH.grid(column=1,row=2)
        YEAR=Entry(self.Add_New_Orders)
        YEAR.grid(column=1,row=3)

        data= pd.read_excel("MOTOR_DATA.xlsx",index_col=0)
        OPTIONS= [data.iloc[i][1] for i in range(0,len(data))]
        MODEL = StringVar(self.Add_New_Orders)
        MODEL.set(OPTIONS[0]) # default value
        w = OptionMenu(self.Add_New_Orders, MODEL, *OPTIONS)
        w.grid(column=1,row=4)

        AMOUNT=Entry(self.Add_New_Orders)
        AMOUNT.grid(column=1,row=5)
        

        self.Quit=Button(self.Add_New_Orders, text="QUIT", fg="red",command=self.master.destroy,width=10,height=2)
        self.Quit.grid(column=1,row=12)

        self.Save=Button(self.Add_New_Orders, text="SAVE", fg="green",command= lambda: Check_and_Add_Orders(DAY.get(),MONTH.get(),YEAR.get(),MODEL.get(),AMOUNT.get()),width=10,height=2)
        self.Save.grid(column=0,row=12)




def Add_Orders_func():
    root=Tk()
    Add_Orders_Application=Add_Orders(master=root)
    Add_Orders_Application.mainloop()


################### Slot_Out ######################
class Slot_Out(Frame):
    def __init__(self,master=None):
        super().__init__(master)
        self.master.configure(background='#f2f6f4')
        self.master.title("DANH SÁCH CÁCH HÃNG XE")
        self.master.geometry("1000x550")
        self.pack()
        self.List_Slot_Out()
    def List_Slot_Out(self):
        self.ListModel=LabelFrame(self.master,text="Danh sách tất cả hãng xe sắp hết",padx=20, pady=20)
        self.ListModel.pack(side='top')
        self.ListView=ttk.Treeview(self.ListModel,height=18)
        self.ListView['columns']=('INDEX','BRAND', "MODEL",'AMOUNT')
    
        self.ListView.column("#0",width=0,minwidth=0)
        self.ListView.column("INDEX",anchor=CENTER,width=50)
        self.ListView.column("BRAND",anchor=CENTER,width=80)
        self.ListView.column("MODEL",anchor=CENTER,width=120)
        self.ListView.column("AMOUNT",anchor=CENTER,width=100)

        self.ListView.heading("INDEX",text="STT",anchor=CENTER)
        self.ListView.heading("BRAND",text="HÃNG",anchor=CENTER)
        self.ListView.heading("MODEL",text="MẪU",anchor=CENTER)
        self.ListView.heading("AMOUNT",text="SỐ LƯỢNG",anchor=CENTER)
        data=pd.read_excel("MOTOR_DATA.xlsx",index_col=0)
        count=0
        for i in range(1,len(data)+1):
            if(data.iloc[i-1][8]<10):
                count=count+1
                self.ListView.insert(parent='', index='end',values=(count,str(data.iloc[i-1][0]),str(data.iloc[i-1][1]),str(data.iloc[i-1][8])))
                
        self.ListView.pack(side=TOP)
        self.Quit=Button(self.ListModel, text="QUIT", fg="red",command=self.master.destroy,width=10,height=2,)
        self.Quit.pack(side=BOTTOM)
        self.Sum=Label(self.ListModel, text="TỔNG CỘNG: "+str(count),padx=5,pady=5)
        self.Sum.pack(side=LEFT)





def Slot_Out_func():
    root=Tk()
    Slot_Out_Application=Slot_Out(master=root)
    Slot_Out_Application.mainloop()

################## Running Out #######################
class Running_Out(Frame):
    def __init__(self,master=None):
        super().__init__(master)
        self.master.configure(background='#f2f6f4')
        self.master.title("DANH SÁCH CÁCH HÃNG XE")
        self.master.geometry("550x550")
        self.pack()
        self.List_Running_Out()
    def List_Running_Out(self):
        self.ListModel=LabelFrame(self.master,text="Danh sách bán chạy",padx=20, pady=20)
        self.ListModel.pack(side='top')
        self.ListView=ttk.Treeview(self.ListModel,height=18)
        self.ListView['columns']=('INDEX', "MODEL",'AMOUNT')
    
        self.ListView.column("#0",width=0,minwidth=0)
        self.ListView.column("INDEX",anchor=CENTER,width=50)
        self.ListView.column("MODEL",anchor=CENTER,width=120)
        self.ListView.column("AMOUNT",anchor=CENTER,width=100)

        self.ListView.heading("INDEX",text="STT",anchor=CENTER)
        self.ListView.heading("MODEL",text="MẪU",anchor=CENTER)
        self.ListView.heading("AMOUNT",text="SỐ LƯỢNG",anchor=CENTER)
        data=pd.read_excel("ORDERS.xlsx",index_col=0)
        
        temp=data.groupby("MODEL").agg({ 'AMOUNT': 'sum'}).reset_index().sort_values(by=['AMOUNT'],ascending=False)
        count=0
        MODEL=list(temp['MODEL'])
        VALUES=list(temp['AMOUNT'])
        for i in range(1,len(temp)+1):
            count=count+1
            self.ListView.insert(parent='', index='end',
            values=(count,MODEL[i-1],VALUES[i-1]))
            
        self.ListView.pack(side=TOP)
        self.Quit=Button(self.ListModel, text="QUIT", fg="red",command=self.master.destroy,width=10,height=2,)
        self.Quit.pack(side=BOTTOM)
        self.Sum=Label(self.ListModel, text="TỔNG CỘNG: "+str(count),padx=5,pady=5)
        self.Sum.pack(side=LEFT)


def Running_Out_func():
    root=Tk()
    Running_Out_Application=Running_Out(master=root)
    Running_Out_Application.mainloop()

########## PROFIT  #######################
def Doanh_Thu():
    data=pd.read_excel("ORDERS.xlsx",index_col=0)
    return str(sum(data["TOTAL"]))
def Tien_Hang():
    data=pd.read_excel("ORDERS.xlsx",index_col=0)
    data_motor=pd.read_excel("MOTOR_DATA.xlsx",index_col=0)
    temp=data.groupby("MODEL").agg({ 'AMOUNT': 'sum'}).reset_index().sort_values(by=['AMOUNT'],ascending=False)
    Sum=0
    for i in range(0,len(temp)):
        Sum=Sum + (data_motor.iloc[list(data_motor.loc[data_motor['TYPE'] == temp.iloc[int(i)]["MODEL"]].index)[0]]['PRICE'])*int(data.iloc[1]['AMOUNT'])
    return Sum
def Loi_Nhuan():
    A=round(float(Doanh_Thu()))
    B=round(float(Tien_Hang()))
    return A-B
class Profit(Frame):
    def __init__(self,master=None):
        super().__init__(master)
        self.master.configure(background='#f2f6f4')
        self.master.title("DOANH THU LỢI NHUẬN")
        self.pack()
        self.List_Profit()
    def List_Profit(self):
        self.ListModel=LabelFrame(self.master,text="Danh sách bán chạy",padx=20, pady=20)
        self.ListModel.pack(side='top')
        L1=Label(self.ListModel,text="Doanh thu: ")
        L1.grid(column=0,row=0)
        L2=Label(self.ListModel, text=Doanh_Thu())
        L2.grid(column=1,row=0)
        L3=Label(self.ListModel,text="Tiền Hàng: ")
        L3.grid(column=0,row=1)
        L4=Label(self.ListModel, text=Tien_Hang())
        L4.grid(column=1,row=1)
        L5=Label(self.ListModel,text="Lợi Nhuận ")
        L5.grid(column=0,row=2)
        L6=Label(self.ListModel, text=round(Loi_Nhuan()))
        L6.grid(column=1,row=2)
        self.Quit=Button(self.ListModel, text="QUIT", fg="red",command=self.master.destroy,width=10,height=2,)
        self.Quit.grid(column=1,row=3)


def Profit_func():
    root=Tk()
    Profit_Application=Profit(master=root)
    Profit_Application.mainloop()