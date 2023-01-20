from tkinter import *
from tkinter import ttk, messagebox, filedialog
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import pickle

#StudentNumber-36222321; Random number generator = 2321+1; random_state= 2322
#instantiating the tkinter root class
main_root = Tk()
main_root.title("Welcome")

#defining the class
class Data:
    #constructor function of the class
    def __init__(self):
        pass   
    
    #function to upload file from local
    def browse_file(self):
        '''function to upload file from local'''
        try:
            filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
            self.df = pd.read_excel(filename)
            self.df = self.df.sample(n=100,random_state=2322)
            messagebox.showinfo(title=None, message="File uploaded and sampling done!")
            process_label = Label(main_root, text= "Press view data to proceed!", font=('Arial', 10), foreground="green")
            process_label.pack(padx=10, pady=10)
        except:
            print("Wrong file format!")
            raise Exception 
    
    #function to display the dataframe created from the excel and perform various operations on it
    def tree_view(self):
        '''function to display the dataframe created from the excel and perform various operations on it'''
        try:
            #defining new window to open on top of the original
            self.root = Toplevel(main_root)
            self.root.geometry("1000x1000")
            self.root.title("Energy Call Centre Analysis")
            #add style
            self.style = ttk.Style()
            self.style.theme_use("default")
            self.style.configure("Treeview", background= "#D3D3D3", foreground = "black", rowheight = 25, fieldbackground="#D3D3D3")
            self.style.map('Treeview', background=[('selected', 'blue')])
            #define tree view for the data
            self.tree_data =ttk.Treeview(self.root)
            self.df_columns = list(self.df.columns)

            #Define our columns
            self.tree_data['columns'] = (self.df_columns[0], self.df_columns[1], self.df_columns[2], self.df_columns[3], self.df_columns[4], self.df_columns[5], self.df_columns[6], self.df_columns[7], self.df_columns[8])

            #Format our columns
            self.tree_data.column("#0", width= 80, minwidth=5)
            for col in self.df_columns:
                self.tree_data.column(col, anchor=W, width=150)

            #Create headings
            self.tree_data.heading("#0", text="Row number", anchor=W)
            for col in self.df_columns:
                self.tree_data.heading(col, text = col, anchor=W)

            #Add data
            global count
            count = 1
            for ind in self.df.index:
                self.tree_data.insert(parent='', index='end', text=count, values=(self.df[self.df_columns[0]][ind], self.df[self.df_columns[1]][ind], self.df[self.df_columns[2]][ind], self.df[self.df_columns[3]][ind], self.df[self.df_columns[4]][ind], self.df[self.df_columns[5]][ind], self.df[self.df_columns[6]][ind], self.df[self.df_columns[7]][ind], self.df[self.df_columns[8]][ind]))
                count += 1

            self.tree_data.pack(pady=20)

            #frame for manipulation
            manipulation_frame = LabelFrame(self.root, text= 'Modify Data')
            manipulation_frame.pack(pady=20)

            #labels for columns
            Label(manipulation_frame, text=self.df_columns[0]).grid(row=0, column=0)
            Label(manipulation_frame, text=self.df_columns[1]).grid(row=0, column=1)
            Label(manipulation_frame, text=self.df_columns[2]).grid(row=0, column=2)
            Label(manipulation_frame, text=self.df_columns[3]).grid(row=0, column=3)
            Label(manipulation_frame, text=self.df_columns[4]).grid(row=0, column=4)
            Label(manipulation_frame, text=self.df_columns[5]).grid(row=0, column=5)
            Label(manipulation_frame, text=self.df_columns[6]).grid(row=0, column=6)
            Label(manipulation_frame, text=self.df_columns[7]).grid(row=0, column=7)
            Label(manipulation_frame, text=self.df_columns[8]).grid(row=0, column=8)

            #input boxes for data
            self.month_box = Entry(manipulation_frame)
            self.vht_box = Entry(manipulation_frame)
            self.tod_box = Entry(manipulation_frame)
            self.agents_box = Entry(manipulation_frame)
            self.callsoff_box = Entry(manipulation_frame)
            self.callsaban_box = Entry(manipulation_frame)
            self.callshand_box = Entry(manipulation_frame)
            self.asa_box = Entry(manipulation_frame)
            self.avghandtime_box = Entry(manipulation_frame)

            self.month_box.grid(row=1, column=0)
            self.vht_box.grid(row=1, column=1)
            self.tod_box.grid(row=1, column=2)
            self.agents_box.grid(row=1, column=3)
            self.callsoff_box.grid(row=1, column=4)
            self.callsaban_box.grid(row=1, column=5)
            self.callshand_box.grid(row=1, column=6)
            self.asa_box.grid(row=1, column=7)
            self.avghandtime_box.grid(row=1, column=8)

            #buttons to perform manipulation in data frame
            add_record = Button(self.root, text='Add Record', command=lambda: self.add_record())
            add_record.pack(padx= 10, pady=10)
        
            remove_record = Button(self.root, text='Delete Record', command=lambda: self.remove_record())
            remove_record.pack(padx= 10, pady=10)

            select_record = Button(self.root, text='Select Record', command=lambda: self.select_record())
            select_record.pack(padx= 10, pady=10)

            update_record = Button(self.root, text='Update Record', command=lambda: self.update_record())
            update_record.pack(padx= 10, pady=10)

            #menu for fetching data based on conditions
            self.options = [
            self.df_columns[0],
            self.df_columns[1],
            self.df_columns[2],
            self.df_columns[3],
            self.df_columns[4],
            self.df_columns[5],
            self.df_columns[6],
            self.df_columns[7],
            self.df_columns[8]
            ]
            #defining combobox to fetch data based on conditions
            self.selectOptions = ttk.Combobox(self.root, values=self.options)
            self.selectOptions.current(0)
            self.selectOptions.bind("<<ComboboxSelected>>", self.selectOptions_click)
            self.selectOptions.pack(padx= 20, pady=20)

            #button to visualise the analysis
            plot = Button(self.root, text='Plot Graphs', command=lambda: self.plot_graph())
            plot.pack(padx= 20, pady=20)
        except:
            print("An exception occurred")
        
    #function to redirect to a different window with different graph options
    def plot_graph(self):
        '''function to redirect to a different window with different graph options'''
        try:
            #defining new window to open on top of the original
            self.plot_window = Toplevel(self.root)
            self.plot_window.geometry("1000x1000")
            self.plot_window.title("Visualizations")
            #defining widgets for the Visualisation window
            vis_label = Label(self.plot_window, text= "Graphical Analysis", font=('Arial', 20))
            vis_label.pack(padx=20, pady=20)
            countplot_frame = LabelFrame(self.plot_window, text= 'Countplot options', width=300, height=500)
            countplot_frame.pack(padx=20, pady=20)
            btn_VHT = Button(countplot_frame, text="Countplot for VHT", command=lambda:self.countplot_VHT())
            btn_VHT.pack(padx=20, pady=20)
            btn_ToD = Button(countplot_frame, text="Countplot for ToD", command=lambda:self.countplot_ToD())
            btn_ToD.pack(padx=20, pady=20)
            scatterplot_frame = LabelFrame(self.plot_window, text= 'Scatterplot options relative to Agents variable', width=300, height=500)
            scatterplot_frame.pack(padx=20, pady=20)
            plot_options = [
                "CallsOffered",
                "CallsHandled",
                "CallsAbandoned"
            ]
            self.select_plotOptions = ttk.Combobox(scatterplot_frame, values=plot_options)
            self.select_plotOptions.current(0)
            self.select_plotOptions.bind("<<ComboboxSelected>>", self.scatterplot_Agents)
            self.select_plotOptions.pack(padx= 20, pady=20)
            btn_scatter = Button(scatterplot_frame, text="Scatterplot for AvgHandle time vs CallsHandled", command=lambda:self.scatter_plot())
            btn_scatter.pack(padx=20, pady=20)
            boxplot_frame = LabelFrame(self.plot_window, text= 'Boxplot options relative to ToD variable', width=300, height=500)
            boxplot_frame.pack(padx=20, pady=20)
            box_options = [
                "CallsOffered",
                "CallsHandled",
                "CallsAbandoned",
                "Avehandletime",
                "ASA"
            ]
            self.select_boxOptions = ttk.Combobox(boxplot_frame, values=box_options)
            self.select_boxOptions.current(0)
            self.select_boxOptions.bind("<<ComboboxSelected>>", self.boxplot_ToD)
            self.select_boxOptions.pack(padx= 20, pady=20)
            barplot_frame = LabelFrame(self.plot_window, text= 'Barplot options relative to VHT variable', width=300, height=500)
            barplot_frame.pack(padx=20, pady=20)
            bar_options = [
                "CallsOffered",
                "CallsHandled",
                "CallsAbandoned",
            ]
            self.select_barOptions = ttk.Combobox(barplot_frame, values=bar_options)
            self.select_barOptions.current(0)
            self.select_barOptions.bind("<<ComboboxSelected>>", self.barplot_VHT)
            self.select_barOptions.pack(padx= 20, pady=20)
            mainplot_frame = LabelFrame(self.plot_window, text= 'Plot to find correlation of all the variables with each other', width=300, height=500)
            mainplot_frame.pack(padx=20, pady=20)
            btn_Heatmap = Button(mainplot_frame, text="HeatMap", command=lambda:self.heatmap())
            btn_Heatmap.pack(padx=20, pady=20)
        except:
            print("An exception occurred")    


    #function to add a new record
    def add_record(self):
        '''function to add a new record'''
        try:
            global count
            #check to validate empty fields
            if self.month_box.get() == "" or self.vht_box.get()=="" or self.tod_box.get()=="" or self.agents_box.get()=="" or self.callsoff_box.get() == "" or self.callsaban_box. get()== "" or self.callshand_box.get()=="" or self.asa_box.get()== "" or self.avghandtime_box.get()=="":
                messagebox.showerror(title=None, message="Empty fields!")
            else: 
                self.tree_data.insert(parent='', index='end', text=count, values=(self.month_box.get(), self.vht_box.get(), self.tod_box.get(), self.agents_box.get(), self.callsoff_box.get(), self.callsaban_box.get(), self.callshand_box.get(), self.asa_box.get(), self.avghandtime_box.get()))
                count +=1
                #Clear the input boxes
                self.month_box.delete(0, END)
                self.vht_box.delete(0, END)
                self.tod_box.delete(0, END)
                self.agents_box.delete(0, END)
                self.callsoff_box.delete(0, END)
                self.callsaban_box.delete(0, END)
                self.callshand_box.delete(0, END)
                self.asa_box.delete(0, END)
                self.avghandtime_box.delete(0, END)
                messagebox.showinfo(title=None, message="Record added successfully!")
        except:
            print("An exception occurred")    
        

    #function to delete selected records
    def remove_record(self):
        '''function to delete selected records'''
        try:
            selected_records = self.tree_data.selection()
            for record in selected_records:
                self.tree_data.delete(record)
            messagebox.showinfo(title=None, message="Record deleted successfully!")    
        except:
            print("An exception occurred")        

    #function to select record
    def select_record(self):
        '''function to select record'''
        try:
            #Clear the input boxes
            self.month_box.delete(0, END)
            self.vht_box.delete(0, END)
            self.tod_box.delete(0, END)
            self.agents_box.delete(0, END)
            self.callsoff_box.delete(0, END)
            self.callsaban_box.delete(0, END)
            self.callshand_box.delete(0, END)
            self.asa_box.delete(0, END)
            self.avghandtime_box.delete(0, END)
            #selecting the data 
            selected = self.tree_data.focus()
            values = self.tree_data.item(selected, 'values')
            #putting selected data to input boxes respectively
            self.month_box.insert(0, values[0])
            self.vht_box.insert(0, values[1])
            self.tod_box.insert(0, values[2])
            self.agents_box.insert(0, values[3])
            self.callsoff_box.insert(0, values[4])
            self.callsaban_box.insert(0, values[5])
            self.callshand_box.insert(0, values[6])
            self.asa_box.insert(0, values[7])
            self.avghandtime_box.insert(0, values[8])
        except:
            print("An exception occurred")    

    #function to update selected record
    def update_record(self):
        '''function to update selected record'''
        try:
            #updating the selected data
            selected = self.tree_data.focus()
            self.tree_data.item(selected, values=(self.month_box.get(), self.vht_box.get(), self.tod_box.get(), self.agents_box.get(), self.callsoff_box.get(), self.callsaban_box.get(), self.callshand_box.get(), self.asa_box.get(), self.avghandtime_box.get()))   
            messagebox.showinfo(title=None, message="Record updated successfully!")
        except:
            print("An exception occurred")

    #function to select menu for fetching data
    def selectOptions_click(self, event):
        '''function to select menu for fetching data'''
        try:
            if self.selectOptions.get() == self.df_columns[0]:
                #defining new window to open on top of the original
                window = Toplevel(self.root)
                window.geometry("400x400")
                #defining widget for the new window
                frame = LabelFrame(window, text= 'Options for the month', width=200, height=300)
                frame.pack(pady=20)
                self.selectMonth = ttk.Combobox(frame, values=["Oct-Nov","Dec-Jan","Feb-Mar"])
                self.selectMonth.current(0)
                self.selectMonth.bind("<<ComboboxSelected>>", self.search_month)
                self.selectMonth.pack(pady=20)
            if self.selectOptions.get() == self.df_columns[1]:
                #defining new window to open on top of the original
                window = Toplevel(self.root)
                window.geometry("400x400")
                frame = LabelFrame(window, text= 'Options for the VHT', width=200, height=300)
                frame.pack(pady=20)
                self.selectVHT = ttk.Combobox(frame, values=["On","Off"])
                self.selectVHT.current(0)
                self.selectVHT.bind("<<ComboboxSelected>>", self.search_VHT)
                self.selectVHT.pack(pady=20)
            if  self.selectOptions.get() == self.df_columns[2]:
                #defining new window to open on top of the original
                window = Toplevel(self.root)
                window.geometry("400x400")
                frame = LabelFrame(window, text= 'Options for time of the day', width=200, height=300)
                frame.pack(pady=20)
                self.selectToD = ttk.Combobox(frame, values=["morning","afternoon","evening"])
                self.selectToD.current(0)
                self.selectToD.bind("<<ComboboxSelected>>", self.search_ToD)
                self.selectToD.pack(pady=20)
            if self.selectOptions.get() == self.df_columns[3]:
                #defining new window to open on top of the original
                window = Toplevel(self.root)
                window.geometry("400x400")
                frame = LabelFrame(window, text= 'Range for agents', width=200, height=100)
                frame.pack(pady=20)
                self.min_label = Label(frame, text="Min")
                self.min_label.pack(pady=10)
                self.Min_Val = Entry(frame, width=20)
                self.Min_Val.pack(pady=5)
                self.max_label = Label(frame, text="Max")
                self.max_label.pack(pady=10)
                self.Max_Val = Entry(frame, width=20)
                self.Max_Val.pack(pady=5)
                self.submit_agent = Button(window, text = "Submit", command= lambda: self.search_Agents())
                self.submit_agent.pack(pady=10)

            if self.selectOptions.get() == self.df_columns[4]:
                #defining new window to open on top of the original
                window = Toplevel(self.root)
                window.geometry("400x400")
                frame = LabelFrame(window, text= 'Range for Calls Offered', width=200, height=100)
                frame.pack(pady=20)
                self.min_label = Label(frame, text="Min")
                self.min_label.pack(pady=10)
                self.Min_Val = Entry(frame, width=20)
                self.Min_Val.pack(pady=5)
                self.max_label = Label(frame, text="Max")
                self.max_label.pack(pady=10)
                self.Max_Val = Entry(frame, width=20)
                self.Max_Val.pack(pady=5)
                self.submit_agent = Button(window, text = "Submit", command= lambda: self.search_CallsOffered())
                self.submit_agent.pack(pady=10)

            if self.selectOptions.get() == self.df_columns[5]:
                #defining new window to open on top of the original
                window = Toplevel(self.root)
                window.geometry("400x400")
                frame = LabelFrame(window, text= 'Range for Calls Abandoned', width=200, height=100)
                frame.pack(pady=20)
                self.min_label = Label(frame, text="Min")
                self.min_label.pack(pady=10)
                self.Min_Val = Entry(frame, width=20)
                self.Min_Val.pack(pady=5)
                self.max_label = Label(frame, text="Max")
                self.max_label.pack(pady=10)
                self.Max_Val = Entry(frame, width=20)
                self.Max_Val.pack(pady=5)
                self.submit_agent = Button(window, text = "Submit", command= lambda: self.search_CallsAbandoned())
                self.submit_agent.pack(pady=10)

            if self.selectOptions.get() == self.df_columns[6]:
                #defining new window to open on top of the original
                window = Toplevel(self.root)
                window.geometry("400x400")
                frame = LabelFrame(window, text= 'Range for Calls Handled', width=200, height=100)
                frame.pack(pady=20)
                self.min_label = Label(frame, text="Min")
                self.min_label.pack(pady=10)
                self.Min_Val = Entry(frame, width=20)
                self.Min_Val.pack(pady=5)
                self.max_label = Label(frame, text="Max")
                self.max_label.pack(pady=10)
                self.Max_Val = Entry(frame, width=20)
                self.Max_Val.pack(pady=5)
                self.submit_agent = Button(window, text = "Submit", command= lambda: self.search_CallsHandled())
                self.submit_agent.pack(pady=10)  

            if self.selectOptions.get() == self.df_columns[7]:
                #defining new window to open on top of the original
                window = Toplevel(self.root)
                window.geometry("400x400")
                frame = LabelFrame(window, text= 'Range for average speed of answer', width=200, height=100)
                frame.pack(pady=20)
                self.min_label = Label(frame, text="Min")
                self.min_label.pack(pady=10)
                self.Min_Val = Entry(frame, width=20)
                self.Min_Val.pack(pady=5)
                self.max_label = Label(frame, text="Max")
                self.max_label.pack(pady=10)
                self.Max_Val = Entry(frame, width=20)
                self.Max_Val.pack(pady=5)
                self.submit_agent = Button(window, text = "Submit", command= lambda: self.search_ASA())
                self.submit_agent.pack(pady=10)

            if self.selectOptions.get() == self.df_columns[8]:
                #defining new window to open on top of the original
                window = Toplevel(self.root)
                window.geometry("400x400")
                frame = LabelFrame(window, text= 'Range for average handle time', width=200, height=100)
                frame.pack(pady=20)
                self.min_label = Label(frame, text="Min")
                self.min_label.pack(pady=10)
                self.Min_Val = Entry(frame, width=20)
                self.Min_Val.pack(pady=5)
                self.max_label = Label(frame, text="Max")
                self.max_label.pack(pady=10)
                self.Max_Val = Entry(frame, width=20)
                self.Max_Val.pack(pady=5)
                self.submit_agent = Button(window, text = "Submit", command= lambda: self.search_AvgHandleTime())
                self.submit_agent.pack(pady=10)
        except:
            print("An exception occurred")              

    #function to display data based on month
    def search_month(self,event):
        '''function to display data based on month'''
        try:
            count=1
            self.tree_data.delete(*self.tree_data.get_children())
            for index, row in self.df.loc[self.df["Month"].eq(self.selectMonth.get())].iterrows():
                self.tree_data.insert("", "end", text =count, values =list(row))
                count+=1
        except:
            print("An exception occurred")        

    #function to display data based on VHT
    def search_VHT(self,event):
        '''function to display data based on VHT'''
        try:
            count=1
            self.tree_data.delete(*self.tree_data.get_children())
            for index, row in self.df.loc[self.df["VHT"].eq(self.selectVHT.get())].iterrows():
                self.tree_data.insert("", "end", text =count, values =list(row))
                count+=1
        except:        
            print("An exception occurred")         

    #function to display data based on ToD
    def search_ToD(self, event):
        '''function to display data based on ToD'''
        try:
            count=1
            self.tree_data.delete(*self.tree_data.get_children())
            for index, row in self.df.loc[self.df["ToD"].eq(self.selectToD.get())].iterrows():
                self.tree_data.insert("", "end", text =count, values =list(row))
                count+=1 
        except:
            print("An exception occurred")         

    #function to display data based on agent range
    def search_Agents(self):
        '''function to display data based on agent range'''
        try:
            count=1
            self.tree_data.delete(*self.tree_data.get_children())
            for index, row in self.df.loc[self.df["Agents"].between(int(self.Min_Val.get()), int(self.Max_Val.get()))].iterrows():
                self.tree_data.insert("", "end", text =count, values =list(row))
                count+=1
        except:
            print("An exception occurred")        

    #function to display data based on Calls offered range
    def search_CallsOffered(self):
        '''function to display data based on Calls offered range'''
        try:
            count=1
            self.tree_data.delete(*self.tree_data.get_children())
            for index, row in self.df.loc[self.df["CallsOffered"].between(int(self.Min_Val.get()), int(self.Max_Val.get()))].iterrows():
                self.tree_data.insert("", "end", text =count, values =list(row))
                count+=1
        except:
            print("An exception occurred")        

    #function to display data based on Calls abandoned range
    def search_CallsAbandoned(self):
        '''function to display data based on Calls abandoned range'''
        try:
            count=1
            self.tree_data.delete(*self.tree_data.get_children())
            for index, row in self.df.loc[self.df["CallsAbandoned"].between(int(self.Min_Val.get()), int(self.Max_Val.get()))].iterrows():
                self.tree_data.insert("", "end", text =count, values =list(row))
                count+=1
        except:
            print("An exception occurred")    

    #function to display data based on Calls handled range
    def search_CallsHandled(self):
        '''function to display data based on Calls handled range'''
        try:
            count=1
            self.tree_data.delete(*self.tree_data.get_children())
            for index, row in self.df.loc[self.df["CallsHandled"].between(int(self.Min_Val.get()), int(self.Max_Val.get()))].iterrows():
                self.tree_data.insert("", "end", text =count, values =list(row))
                count+=1
        except:
            print("An exception occurred")    

    #function to display data based on average speed of answer range
    def search_ASA(self):
        '''function to display data based on average speed of answer range'''
        try:
            count=1
            self.tree_data.delete(*self.tree_data.get_children())
            for index, row in self.df.loc[self.df["ASA"].between(float(self.Min_Val.get()), float(self.Max_Val.get()))].iterrows():
                self.tree_data.insert("", "end", text =count, values =list(row))
                count += 1
        except:
            print("An exception occurred")        

    #function to display data based on average handle time range
    def search_AvgHandleTime(self):
        '''function to display data based on average handle time range'''
        try:
            count=1
            self.tree_data.delete(*self.tree_data.get_children())
            for index, row in self.df.loc[self.df["Avehandletime"].between(float(self.Min_Val.get()), float(self.Max_Val.get()))].iterrows():
                self.tree_data.insert("", "end", text =count, values =list(row))
                count += 1
        except:
            print("An exception occurred")         

    #function to plot the countplot of VHT
    def countplot_VHT(self):
        '''function to plot the countplot of VHT'''
        try:
            plot_root = Toplevel(self.plot_window)
            plot_root.geometry("1000x1000")
            frame = LabelFrame(plot_root, text= 'Countplot', width=500, height=500)
            frame.pack(pady=20)
            fig = Figure(figsize=(70,70), dpi=100)
            a_subplot = fig.add_subplot(111)
            sns.countplot(x=self.df['VHT'],data=self.df, ax= a_subplot)
            canvas = FigureCanvasTkAgg(fig, frame)
            canvas.get_tk_widget().pack()
            canvas.draw()
        except:
            print("An exception occurred")

    #function to plot the countplot of ToD
    def countplot_ToD(self):
        '''function to plot the countplot of ToD'''
        try:
            plot_root = Toplevel(self.plot_window)
            plot_root.geometry("1000x1000")
            frame = LabelFrame(plot_root, text= 'Countplot', width=500, height=500)
            frame.pack(pady=20)
            fig = Figure(figsize=(70,70), dpi=100)
            a_subplot = fig.add_subplot(111)
            sns.countplot(x=self.df['ToD'],data=self.df, ax= a_subplot)
            canvas = FigureCanvasTkAgg(fig, frame)
            canvas.get_tk_widget().pack()
            canvas.draw()
        except:
            print("An exception occurred")

    #function to plot scatterplot relative to Agents
    def scatterplot_Agents(self, event):
        '''function to plot scatterplot relative to Agents'''
        try:
            plot_root = Toplevel(self.plot_window)
            plot_root.geometry("1500x1500")
            frame = LabelFrame(plot_root, text= 'Scatterplot', width=500, height=500)
            frame.pack(pady=20)
            fig = Figure(figsize=(100,100), dpi=100)
            a_subplot = fig.add_subplot(111)
            if self.select_plotOptions.get() == "CallsOffered":
                sns.scatterplot(data=self.df, x=self.df['Agents'], y=self.df['CallsOffered'], hue=self.df['VHT'], ax=a_subplot)
            if self.select_plotOptions.get() == "CallsAbandoned":
                sns.scatterplot(data=self.df, x=self.df['Agents'], y=self.df['CallsAbandoned'], hue=self.df['VHT'], ax=a_subplot)
            if self.select_plotOptions.get() == "CallsHandled":
                sns.scatterplot(data=self.df, x=self.df['Agents'], y=self.df['CallsHandled'], hue=self.df['VHT'], ax=a_subplot)        
            canvas = FigureCanvasTkAgg(fig, frame)
            canvas.get_tk_widget().pack()
            canvas.draw()
        except:
            print("An exception occurred")    

    #function to plot scatterplot between Avghandletime and CallsHandled
    def scatter_plot(self):
        '''function to plot scatterplot between Avghandletime and CallsHandled'''
        try:
            plot_root = Toplevel(self.plot_window)
            plot_root.geometry("1500x1500")
            frame = LabelFrame(plot_root, text= 'Scatterplot', width=500, height=500)
            frame.pack(pady=20)
            fig = Figure(figsize=(100,100), dpi=100)
            a_subplot = fig.add_subplot(111)
            sns.scatterplot(data=self.df, x=self.df['Avehandletime'], y=self.df['CallsHandled'], hue=self.df['VHT'], ax=a_subplot)
            canvas = FigureCanvasTkAgg(fig, frame)
            canvas.get_tk_widget().pack()
            canvas.draw()
        except:
            print("An exception occurred")    

    #function to plot boxplot relative to ToD
    def boxplot_ToD(self, event):
        '''function to plot boxplot relative to ToD'''
        try:
            plot_root = Toplevel(self.plot_window)
            plot_root.geometry("1500x1500")
            frame = LabelFrame(plot_root, text= 'Boxplot', width=500, height=500)
            frame.pack(pady=20)
            fig = Figure(figsize=(100,100), dpi=100)
            a_subplot = fig.add_subplot(111)
            if self.select_boxOptions.get() == "CallsOffered":
                sns.boxplot(x=self.df['ToD'], y=self.df['CallsOffered'],data=self.df, palette='rainbow', hue=self.df['VHT'], ax=a_subplot)
            if self.select_boxOptions.get() == "CallsHandled":
                sns.boxplot(x=self.df['ToD'], y=self.df['CallsHandled'],data=self.df, palette='rainbow', hue=self.df['VHT'], ax=a_subplot)
            if self.select_boxOptions.get() == "CallsAbandoned":
                sns.boxplot(x=self.df['ToD'], y=self.df['CallsAbandoned'],data=self.df, palette='rainbow', hue=self.df['VHT'], ax=a_subplot)
            if self.select_boxOptions.get() == "Avehandletime":
                sns.boxplot(x=self.df['ToD'], y=self.df['Avehandletime'],data=self.df, palette='rainbow', hue=self.df['VHT'], ax=a_subplot)
            if self.select_boxOptions.get() == "ASA":
                sns.boxplot(x=self.df['ToD'], y=self.df['ASA'],data=self.df, palette='rainbow', hue=self.df['VHT'], ax=a_subplot)
            canvas = FigureCanvasTkAgg(fig, frame)
            canvas.get_tk_widget().pack()
            canvas.draw()     
        except:
            print("An exception occurred")    

    #function to plot the heatmap
    def heatmap(self):
        '''function to plot the heatmap'''
        try:
            plot_root = Toplevel(self.plot_window)
            plot_root.geometry("1500x1500")
            frame = LabelFrame(plot_root, text= 'Heatmap', width=500, height=500)
            frame.pack(pady=20)
            fig = Figure(figsize=(100,100), dpi=100)
            a_subplot = fig.add_subplot(111)
            corr_var=self.df.corr()
            sns.heatmap(corr_var, square=True, cbar=False, ax=a_subplot, annot= True)
            canvas = FigureCanvasTkAgg(fig, frame)
            canvas.get_tk_widget().pack()
            canvas.draw()
        except:
            print("An exception occurred")    

    #function to plot the baroplot relative to VHT
    def barplot_VHT(self, event):
        '''function to plot the baroplot relative to VHT'''
        try:
            plot_root = Toplevel(self.plot_window)
            plot_root.geometry("1000x1000")
            frame = LabelFrame(plot_root, text= 'Barplot', width=500, height=500)
            frame.pack(pady=20)
            fig = Figure(figsize=(100,100), dpi=100)
            a_subplot = fig.add_subplot(111)
            if self.select_barOptions.get() == "CallsOffered":
                sns.barplot(data=self.df, x=self.df['VHT'], y=self.df['CallsOffered'], hue=self.df['ToD'], ax=a_subplot)
            if self.select_barOptions.get() == "CallsAbandoned":
                sns.barplot(data=self.df, x=self.df['VHT'], y=self.df['CallsAbandoned'], hue=self.df['ToD'], ax=a_subplot)
            if self.select_barOptions.get() == "CallsHandled":
                sns.barplot(data=self.df, x=self.df['VHT'], y=self.df['CallsHandled'], hue=self.df['ToD'], ax=a_subplot)
            canvas = FigureCanvasTkAgg(fig, frame)
            canvas.get_tk_widget().pack()
            canvas.draw()
        except:
            print("An exception occurred")     

#instantiating the class    
view = Data()
#checking if there already exists pickle file
try :
    file = open("record.ms","rb")
    prev_run = pickle.load(file)
    print("Previous analysis", prev_run)
    file.close()
except FileNotFoundError :
    #file doesn't exist
    print("A new file is created")
#new file is created to store the current state in serialized form   
file = open("record.ms","wb")
pickle.dump(view,file)
file.close()

#defining the widgets of the main window here
welcome_label = Label(main_root, text= "Welcome User!", font=('Arial', 20))
welcome_label.pack(padx=20, pady=20)
upload_file = Button(main_root, text="Upload file", command=lambda: view.browse_file())
upload_file.pack(padx=10, pady=10)
view_data = Button(main_root, text = "View Data", command= lambda: view.tree_view())
view_data.pack(padx=10, pady=10)        
main_root.geometry("600x600")
main_root.mainloop()