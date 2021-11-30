from tkinter import *
from data import *
from tkinter import ttk
import numpy as np
################# APPLICATION #################
class Application(Frame):
    def __init__(self,master=None):
        super().__init__(master)
        self.master.configure(background='#f2f6f4')
        self.master.title("HỆ THỐNG QUẢN LÝ XE MÁY ")
        self.master.geometry("450x400")
        self.pack()
        self.Crate_Menu()

    def Crate_Menu(self):
        self.manage_frame=ttk.Notebook(self)
        self.manage_frame.pack(pady=15)
        
        Manage=Frame(self.manage_frame,width=600,height=400)
        Orders=Frame(self.manage_frame,width=600,height=400)
        Analyze=Frame(self.manage_frame,width=600,height=400)

        Manage.grid(sticky="nsew")
        Orders.grid(sticky="nsew")
        Analyze.grid(sticky="nsew")
        
        self.manage_frame.add(Manage,text="Danh Mục")

        List_Model=Button(Manage,text='Xem danh sách sản phẩm',command=List_Model_func).grid(row=0,column=0,pady=10,sticky="nsew")
        Add_Model=Button(Manage,text='Thêm sản phẩm',command=Add_Model_func).grid(row=1,column=0,pady=10,sticky="nsew")
        Del_Model=Button(Manage,text='Xóa sản phẩm',command=Del_Model_func).grid(row=2,column=0,pady=10,sticky="nsew")
        Config_Model=Button(Manage,text='Sửa sản phẩm',command=Edit_Model_func).grid(row=3,column=0,pady=10,sticky="nsew")

        List_Brand=Button(Manage,text='Xem danh sách các hãng xe',command=List_Brand_func).grid(row=0,column=1,pady=10,sticky="nsew")
        Add_Brand=Button(Manage,text='Thêm hãng xe',command=Add_Model_func).grid(row=1,column=1,pady=10,sticky="nsew")
        Del_Brand=Button(Manage,text='Xóa hãng xe',command=Del_Brand_func).grid(row=2,column=1,pady=10,sticky="nsew")
        Config_Brand=Button(Manage,text='Sửa hãng xe',command=Edit_Brand_func).grid(row=3,column=1,pady=10,sticky="nsew")

        self.manage_frame.add(Orders,text="Đơn Hàng")
        List_Orders_Label=Label(Orders,text="Chọn Tháng: ").grid(row=0,column=0) 
        data_orders= pd.read_excel("ORDERS.xlsx",index_col=0)
        OPTIONS= list(data_orders["MONTH"].value_counts().keys())
        OPTIONS=np.sort(OPTIONS)
        variable1 = StringVar(Orders)
        variable1.set(OPTIONS[0]) # default value
        w = OptionMenu(Orders, variable1, *OPTIONS)
        w.grid(column=1,row=0)
        List_Orders=Button(Orders,text='Xem Danh Sách Đơn Hàng',comman=lambda: Orders_Info_func(variable1.get())).grid(row=1,column=1,pady=10,sticky="nsew")
        Orders_Details=Button(Orders,text='Tìm kiếm thông tin đơn hàng',command=Find_Orders_Info_func).grid(row=2,column=1,pady=10,sticky="nsew")
        Config_Orders=Button(Orders,text='Chỉnh sửa thông tin đơn hàng',command=Edit_Orders_Info_func).grid(row=3,column=1,pady=10,sticky="nsew")
        Del_Orders=Button(Orders,text='Xóa đơn hàng',command=Del_Orders_func).grid(row=4,column=1,pady=10,sticky="nsew")
        Add_Orders=Button(Orders,text='Thêm đơn hàng',command=Add_Orders_func).grid(row=5,column=1,pady=10,sticky="nsew")

        self.manage_frame.add(Analyze,text="Thống Kê")
        
        Running_Out=Button(Analyze,text='Mặt hàng sắp hết',command=Slot_Out_func).grid(row=0,column=1,pady=10,sticky="nsew")
        Best_Selling=Button(Analyze,text='Mặt hàng bán chạy',command=Running_Out_func).grid(row=1,column=1,pady=10,sticky="nsew")
        Revenue_Profit=Button(Analyze,text='Doanh thu và lợi nhuận',command=Profit_func).grid(row=2,column=1,pady=10,sticky="nsew")

        Exit=Button(self,text='QUIT',width=10,height=10,background='#faeee7',fg="red",command=self.master.destroy).pack(pady=10)

def MainApplication():
    root = Tk()
    app = Application(master=root)
    app.mainloop()

MainApplication()