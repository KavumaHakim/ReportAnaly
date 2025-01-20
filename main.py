#!/usr/bin/python3
from datetime import date
from textwrap import fill
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pandas as pd
import tkinter as tk
from tkinter import filedialog 
from CTkMessagebox import CTkMessagebox
import tkinter.ttk as ttk
from tkinter import Tk
from CTkTable import CTkTable
import CTkToolTip
from customtkinter import (
    CTk,
    CTkButton,
    CTkEntry,
    CTkFont,
    CTkProgressBar,
    CTkFrame,
    CTkLabel,
    CTkRadioButton,
    CTkSwitch,
    CTkTabview,
    set_appearance_mode,
    set_default_color_theme)
from uri_template import expand


class splash:
    def __init__(self):
        # build ui
        self.win = Tk()
        self.win.configure(background="black")
        self.win.configure(relief="flat")
        self.win.geometry("600x400+500+150")
        self.win.overrideredirect(True)
        self.win.resizable(False, False)
        self.title = ttk.Label(self.win)
        self.title.configure(background="black",font="{Elephant} 30 {bold italic}",foreground="#db0967",justify="center",state="normal",text='Simple Report Analysis\nProgram\nVersion 2.0\nHaKiM')
        self.title.pack(pady=30, side="top")
        self.progress = CTkProgressBar(self.win, orientation="horizontal")
        self.load = tk.IntVar()
        self.progress.configure(border_color="#d93ca2",border_width=2,corner_radius=30,height=75,mode="determinate",progress_color="#db0967",variable=self.load,width=500)
        self.progress.pack(side="top")
        self._progress()
        
        self.win.mainloop()
    def _progress(self):
        print(self.progress.get())
        if self.progress.get() < .98:
            self.progress.step()
            self.win.after(50, self._progress)
        elif self.progress.get() >= .9:
            self.progress.stop()
            self.win.destroy()

class ReportApp:
    def __init__(self, master=None):
        # build ui
        self.root = CTk(None)
        self.root.configure(cursor="arrow", width=2)
        set_appearance_mode("dark")
        set_default_color_theme("blue")
        self.root.iconbitmap("icon2.ico")
        self.root.minsize(665, 595)
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.title("PERFORMANCE")
        self.TabV = CTkTabview(self.root)
        
        ###############  View frame ####################
        View_tab = self.TabV.add("    View    ")
        self.View_Frame = CTkFrame(View_tab,corner_radius=0)
        ctkframe_14 = CTkFrame(self.View_Frame,border_width=1, corner_radius=15)
        self.bio_data_fr = CTkFrame(ctkframe_14,border_width=1,corner_radius=15,fg_color="#2b2b2b")
        self.label_1 = CTkLabel(self.bio_data_fr, text="Name  :", font=("century gothic", 19,"bold"), text_color="#db0967")
        self.label_1.grid(column=0, columnspan=2, padx=10, pady=20, row=0)
        self.Name_display = CTkLabel(self.bio_data_fr, font=("century gothic", 19,"bold"),text='_______________', text_color="#db0967")
        self.Name_display.grid(column=2, columnspan=3, pady=20, row=0)
        self.label_5 = CTkLabel(self.bio_data_fr, text="Gender  :", font=("century gothic", 19,"bold"), text_color="#db0967")
        self.label_5.grid(column=0, columnspan=2, padx=10, pady=15, row=1)
        self.Gender_display = CTkLabel(self.bio_data_fr, font=("century gothic", 19,"bold"),text='_______________', text_color="#db0967")
        self.Gender_display.grid(column=2, columnspan=3, pady=15, row=1)
        self.label_9 = CTkLabel(self.bio_data_fr, text="Date  :", font=("century gothic", 19,"bold"), text_color="#db0967")
        self.label_9.grid(column=0, columnspan=2, padx=10, pady=20, row=3)
        self.Date_today_display = CTkLabel(self.bio_data_fr, font=("century gothic", 19,"bold"),text='_______________', text_color="#db0967")
        self.Date_today_display.grid(column=2, columnspan=3, pady=15, row=3)
        self.bio_data_fr.pack(expand=True,fill="both",padx=5,pady=5,side="left")
        self.results_frame = CTkFrame(ctkframe_14,border_width=1,corner_radius=15,fg_color="#2b2b2b")
        self.label_11 = ttk.Label(self.results_frame, background="#2b2b2b",cursor="arrow",font="{century gothic} 14 {bold}",foreground="#db0967",justify="left",text='Best Subject: ')
        self.label_11.grid(column=0, padx=20, pady=20, row=0)
        self.Best_subject_display = ttk.Label(self.results_frame, background="#2b2b2b",cursor="arrow",font="{century gothic} 16 {bold underline}",foreground="#db0967",text='Some Subject')
        self.Best_subject_display.grid(column=1, columnspan=2, padx=10, pady=20, row=0)
        self.label_13 = ttk.Label(self.results_frame, background="#2b2b2b",cursor="arrow",font="{century gothic} 14  {bold}",foreground="#db0967",justify="left",text='Worst Subject: ')
        self.label_13.grid(column=0, padx=20, pady=20, row=1)
        self.worst_subject_display = ttk.Label(self.results_frame, background="#2b2b2b",cursor="arrow",font="{century gothic} 16 {bold italic}",foreground="#db0967",text='Some Subject')
        self.worst_subject_display.grid(column=1, columnspan=2, padx=10, pady=20, row=1)
        self.label_15 = ttk.Label(self.results_frame,background="#2b2b2b", font="{century gothic} 14 {bold}",foreground="#db0967",justify="left",text='Average : ')
        self.label_15.grid(column=0, padx=20, pady=20, row=2)
        self.Average_score_display = ttk.Label(self.results_frame, background="#2b2b2b",cursor="arrow",font="{century gothic} 16 {bold}",foreground="#db0967",text='99.99')
        self.Average_score_display.grid(column=1, columnspan=2, padx=10, pady=20, row=2)
        self.results_frame.pack(expand=True,fill="both",padx=5,pady=5,side="right")
        ctkframe_14.pack(anchor="n",expand=True,fill="both",padx=5,pady=5,side="top")
        self.table_frame = CTkFrame(self.View_Frame, corner_radius=25)

        self.Table = CTkTable(self.table_frame,row=10, column=2,hover_color="#db0967", hover=True, header_color="#db0967", border_color="#db0967",corner_radius=25,)
        
        self.table_frame.pack(anchor="s",expand=True,fill="both",padx=5,pady=5,side="bottom")
        self.View_Frame.pack(expand=True, fill="both", side="top")

        ###############  Graphical frame ####################

        Graphical_tab = self.TabV.add("    Graphical    ")
        self.Graphical_Frame = CTkFrame(Graphical_tab)
        self.div1 = CTkFrame(master=self.Graphical_Frame, bg_color="transparent", border_color="red", border_width=1)
        self.div2 = CTkFrame(master=self.Graphical_Frame, bg_color="transparent", border_color="red", border_width=1)

        self.graph_1 = CTkFrame(master=self.div1, bg_color="transparent", border_color="red", border_width=1)
        self.graph_2 = CTkFrame(master=self.div1, bg_color="transparent", border_color="red", border_width=1)

        self.graph_3 = CTkFrame(master=self.div2, bg_color="transparent", border_color="red", border_width=1)
        self.graph_4 = CTkFrame(master=self.div2, bg_color="transparent", border_color="red", border_width=1)

        self.div1.pack(fill="both", side="top", anchor="n", expand=True)
        self.div2.pack(fill="both", side="bottom", anchor="s",  expand=True)

        self.graph_1.pack(side="left", anchor="s", fill="both", expand=True)
        self.graph_2.pack( side="right", anchor="s", fill="both", expand=True)

        self.graph_3.pack(side="left", anchor="s", fill="both", expand=True)
        self.graph_4.pack( side="right", anchor="s", fill="both", expand=True)
        
        self.Graphical_Frame.pack(expand=True, fill="both", side="top")







        ###############  Data collection frame ####################
        self.scores = {
            "Subjects" : ["Mathematics", "English","Geography", "RE","History", "Physics","Biology", "Chemistry","Elective"],
            "Scores": [],
            }
    
        Data_tab = self.TabV.add("    Data    ")
        self.Data_Frame = CTkFrame(Data_tab)
        self.Enter_Data_Title = CTkLabel(self.Data_Frame,font=CTkFont("Arial Black",25,"bold","italic",False,False),justify="left",text='ENTER DATA',underline=0)
        self.Enter_Data_Title.pack(side="top")
        self.Manual_Entry = CTkFrame(self.Data_Frame,corner_radius=15)
        self.Name_label = CTkLabel(self.Manual_Entry,font=CTkFont("Century Gothic",24,"normal","roman",False,False),text='Name')
        self.Name_label.pack(side="top", pady=2)
        self.Name_entry = CTkEntry(self.Manual_Entry,font=CTkFont("Century Gothic",16,"normal","roman",False,False),placeholder_text="                First-Name   Last-Name")
        self.Name_entry.pack(fill="x", padx=7, pady=5, side="top")
        self.Subject_scores_title = CTkLabel(self.Manual_Entry,font=CTkFont("Century Gothic",20,"normal","roman",False,False),text='Subject scores')
        self.Subject_scores_title.pack(side="top")
        self.mtc_data_frame = CTkFrame(self.Manual_Entry)
        self.mtc = CTkLabel(self.mtc_data_frame,text='Mathematics')
        self.mtc.pack(expand=True, side="left", pady=3)
        self.mtc_s = CTkEntry(self.mtc_data_frame)
        self.tip = CTkToolTip.CTkToolTip(widget=self.mtc_s,message="You can select the next Entry by pressing Tab on the keyboard",
                                        alpha=1,border_color="#db0967",border_width=2,delay=0.5)
        self.mtc_s.pack(side="right", padx=5)
        self.mtc_data_frame.pack(expand=False,fill="x",padx=5,pady=1,side="top")
        self.eng_data_frame = CTkFrame(self.Manual_Entry)
        self.eng = CTkLabel(self.eng_data_frame, text='English')
        self.eng.pack(expand=True, side="left", pady=3)
        self.eng_s = CTkEntry(self.eng_data_frame)
        self.eng_s.pack(side="right", padx=5)
        self.eng_data_frame.pack(expand=False,fill="x",padx=5,pady=1,side="top")
        self.geog_data_frame = CTkFrame(self.Manual_Entry)
        self.geog = CTkLabel(self.geog_data_frame,text='Geography')
        self.geog.pack(expand=True, side="left", pady=3)
        self.geog_s = CTkEntry(self.geog_data_frame)
        self.geog_s.pack(side="right", padx=5)
        self.geog_data_frame.pack(expand=False,fill="x",padx=5,pady=1,side="top")
        self.re_data_frame = CTkFrame(self.Manual_Entry)
        self.re = CTkLabel(self.re_data_frame, text='Religion')
        self.re.pack(expand=True, side="left", pady=3)
        self.re_s = CTkEntry(self.re_data_frame)
        self.re_s.pack(side="right", padx=5)
        self.re_data_frame.pack(expand=False,fill="x",padx=5,pady=1,side="top")
        self.hist_data_frame = CTkFrame(self.Manual_Entry)
        self.hist = CTkLabel(self.hist_data_frame, text='History')
        self.hist.pack(expand=True, side="left", pady=3)
        self.hist_s = CTkEntry(self.hist_data_frame)
        self.hist_s.pack(side="right", padx=5)
        self.hist_data_frame.pack(expand=False,fill="x",padx=5,pady=2,side="top")
        self.phy_data_frame = CTkFrame(self.Manual_Entry)
        self.phy = CTkLabel(self.phy_data_frame,text='Physics')
        self.phy.pack(expand=True, side="left", pady=3)
        self.phy_s = CTkEntry(self.phy_data_frame)
        self.phy_s.pack(side="right", padx=5)
        self.phy_data_frame.pack(expand=False,fill="x",padx=5,pady=1,side="top")
        self.bio_data_frame = CTkFrame(self.Manual_Entry)
        self.bio = CTkLabel(self.bio_data_frame, text='Biology')
        self.bio.pack(expand=True, side="left", pady=3)
        self.bio_s = CTkEntry(self.bio_data_frame)
        self.bio_s.pack(side="right", padx=5)
        self.bio_data_frame.pack(expand=False,fill="x",padx=5,pady=1,side="top")
        self.chem_data_frame = CTkFrame(self.Manual_Entry)
        self.chem = CTkLabel(self.chem_data_frame,text='Chemistry')
        self.chem.pack(expand=True, side="left", pady=3)
        self.chem_s = CTkEntry(self.chem_data_frame)
        self.chem_s.pack(side="right", padx=5)
        self.chem_data_frame.pack(expand=False,fill="x",padx=5,pady=2,side="top")
        self.elect_data_frame = CTkFrame(self.Manual_Entry)
        self.elect = CTkLabel(self.elect_data_frame,text='Elective Subject')
        self.elect.pack(expand=True, side="left", pady=3)
        self.elect_s = CTkEntry(self.elect_data_frame)
        self.elect_s.pack(side="right", padx=5)
        self.elect_data_frame.pack(expand=False, fill="x", padx=5, pady=1, side="top")

        self.radio_var = tk.IntVar()

        self.male = CTkRadioButton(self.Manual_Entry, value=1, variable=self.radio_var,hover_color="#db0967", text='Male')
        self.male.pack(anchor="n", padx=10, pady=5, side="left")

        self.female = CTkRadioButton(self.Manual_Entry, value=2, variable=self.radio_var, hover_color="#db0967", text='Female')
        self.female.pack(anchor="n", padx=10, pady=5, side="right")

        self.submit_button = CTkButton(self.Manual_Entry, hover=True, command=self.submit)
        self.submit_button.configure(corner_radius=15,fg_color="#db0967",font=CTkFont("georgia",16,"bold","roman",False,False),hover_color="#b15ee3",text='Submit')
        self.submit_button.pack(anchor="s",expand=True,fill="x",side="bottom")

        self.Manual_Entry.pack(anchor="nw",expand=True,fill="both",padx=5,pady=5,side="left")

        self.path_entry = CTkFrame(self.Data_Frame,corner_radius=15)

        self.Import = CTkLabel(self.path_entry)
        self.Import.configure(font=CTkFont("century gothic",24,"normal","roman",False,False),text='Import from File')
        self.Import.pack(pady=5, side="top")

        ctklabel_17 = CTkLabel(self.path_entry,text='From Excel file')
        ctklabel_17.pack(side="top")

        separator_1 = ttk.Separator(self.path_entry,orient="horizontal")
        separator_1.pack(fill="x", padx=15, side="top")

        ctkframe_1 = CTkFrame(self.path_entry)

        self.file_path = CTkEntry(ctkframe_1,corner_radius=0,placeholder_text="File path")
        self.file_path.pack(padx=0, side="left")

        self.browse_files = CTkButton(ctkframe_1, hover=True,command= self.import_file)
        self.browse_files.configure(corner_radius=0,fg_color="#db0967",font=CTkFont("georgia",16,"normal","roman",False,False),hover_color="#b15ee3",text='Browse',width=5)
        self.browse_files.pack(padx=0, side="top")

        ctkframe_1.pack(pady=10, side="top")

        self.switch = CTkSwitch(self.path_entry,hover=True,text_color="#ffffff",command=self.toggled_mode)
        self.switch.configure(fg_color="#db0967",font=CTkFont("CENTURY GOTHIC",24,"normal","roman",False,False),state="normal",text='Light Mode')
        self.switch.pack(anchor="center", pady=10, side="bottom")
        
        self.path_entry.pack(anchor="n",expand=True,fill="y",ipadx=30,padx=5,pady=5,side="right")
        self.Data_Frame.pack(expand=True, fill="both", side="top")

        self.large_data_button = CTkButton(master=self.path_entry, text="Large data", cursor='hand2', hover_color="#db0967", command=self.large_data)
        self.large_data_button.pack(expand=True,fill="x", padx=10, anchor="n")

        self.TabV.pack(expand=True, fill="both", side="top")

        # Main widget
        self.mainwindow = self.root
    
    def run(self):
        self.mainwindow.mainloop()


    def large_data(self):

        self.data = tk.Toplevel(self.root, takefocus=1)
        self.data.overrideredirect(True)
        self.data.grab_set()
        self.data.focus_set()

        self.large_data_Frame = CTkFrame(self.data)
        self.ctkframe_2 = CTkFrame(self.large_data_Frame)
        self.ctkframe_2.configure(
            border_color="#eb56eb",
            border_width=5,
            height=200)
        self.ctkframe_3 = CTkFrame(self.ctkframe_2)
        self.ctkframe_3.configure(border_color="#eb56eb", border_width=1, width=300)
        self.ctklabel_24 = CTkLabel(self.ctkframe_3)
        self.ctklabel_24.configure(
            anchor="center",
            cursor="circle",
            justify="center",
            text='BRIEF OVERVIEW')
        self.ctklabel_24.pack(anchor="center", padx=20, pady=5, side="top")
        self.average = CTkLabel(self.ctkframe_3, text='Average: ')
        self.average.pack(anchor="w", expand=False, padx=5, pady=0, side="top")
        self.Students = CTkLabel(self.ctkframe_3,text='Number of Students: ')
        self.Students.pack(anchor="w", expand=False, padx=5, side="top")
        self.grade = CTkLabel(self.ctkframe_3,text='Average Grade: ')
        self.grade.pack(anchor="w", padx=5, side="top")
        self.avg_boys = CTkLabel(self.ctkframe_3, text='Average Boys:')
        self.avg_boys.pack(anchor="w", padx=5, side="top")
        self.avg_girls = CTkLabel(self.ctkframe_3, text='Average Girls:')
        self.avg_girls.pack(anchor="w", padx=5, side="top")
        self.ctkframe_3.pack(expand=True, fill="both", side="left")
        self.ctkframe_3.pack_propagate(False)
        self.ctkframe_4 = CTkFrame(self.ctkframe_2)
        self.ctkframe_4.configure(border_color="#eb56eb", border_width=1, width=300)
        self.ctklabel_33 = CTkLabel(self.ctkframe_4,cursor="circle", text='OTHER INFO')
        self.ctklabel_33.pack(anchor="center", pady=5, side="top")
        self.BEST = CTkLabel(self.ctkframe_4,text='BEST STUDENT: ')
        self.BEST.pack(anchor="w", padx=5, side="top")
        self.Worst = CTkLabel(self.ctkframe_4,text='LAST STUDENT:')
        self.Worst.pack(anchor="w", padx=5, side="top")
        self.BEST_SUB = CTkLabel(self.ctkframe_4, text='BEST SUBJECT:')
        self.BEST_SUB.pack(anchor="w", padx=5, side="top")
        self.WORST_SUB = CTkLabel(self.ctkframe_4,text='WORST SUBJECT: ')
        self.WORST_SUB.pack(anchor="w", padx=5, side="top")
        self.ctkframe_4.pack(expand=True, fill="both", side="right")
        self.ctkframe_4.pack_propagate(False)
        self.ctkframe_2.pack(side="top")
        self.ctkframe_5 = CTkFrame(self.large_data_Frame)
        ctkbutton_1 = CTkButton(self.ctkframe_5, hover=True)
        ctkbutton_1.configure(hover_color="#ff0080", text='GENERATE REPORTS')
        ctkbutton_1.pack(
            anchor="n",
            expand=True,
            fill="both",
            pady=5,
            side="top")
        self.CHARTS = CTkButton(self.ctkframe_5, hover=True)
        self.CHARTS.configure(hover_color="#ff0080", text='DRAW CHARTS')
        self.CHARTS.pack(
            anchor="n",
            expand=True,
            fill="both",
            pady=0,
            side="top")
        self.exit = CTkButton(self.ctkframe_5, text="Exit", command=lambda : self.data.destroy(), fg_color="crimson", hover_color="green")
        self.exit.pack(fill="x")
        self.ctkframe_5.pack(expand=True, fill="both", side="top")
        self.ctkframe_5.pack_propagate(False)
        self.large_data_Frame.pack(expand=True, fill="both", side="top")



    def toggled_mode(self):
        if self.switch.get() == 1 :
            set_appearance_mode("light")
            self.switch.configure(text="Dark mode")
            self.results_frame.configure(fg_color="#dbdbdb")
            self.bio_data_fr.configure(fg_color="#dbdbdb")
            for i in [self.label_11,
                    self.label_13, self.label_15,
                    self.Best_subject_display,self.worst_subject_display,
                    self.Average_score_display]:
                i.configure(background="#dbdbdb")
            
        elif self.switch.get() == 0:
            set_appearance_mode("dark")
            self.switch.configure(text="Light mode")
            self.results_frame.configure(fg_color="#2b2b2b")
            self.bio_data_fr.configure(fg_color="#2b2b2b")
            for i in [self.label_11,
                    self.label_13, self.label_15,
                    self.Best_subject_display,self.worst_subject_display,
                    self.Average_score_display]:
                i.configure(background="#2b2b2b")

    def submit(self):
        try:
            self.scores["Scores"] = [
                int(self.mtc_s.get()),
                int(self.eng_s.get()),
                int(self.geog_s.get()),
                int(self.re_s.get()),
                int(self.hist_s.get()),
                int(self.phy_s.get()),
                int(self.bio_s.get()),
                int(self.chem_s.get()),
                int(self.elect_s.get())
                ]
        except:
            CTkMessagebox(master=self.Data_Frame,
                        title="Error", icon="warning",
                        message="Error while entering data!\nPlease check the values\nEnsure that all values are correctly typed",
                        icon_size=(50,50),
                        corner_radius=25,
                        sound=True,
                        font=("century gothic",15, "bold"),
                        header=False,
                        topmost=True,
                        text_color="#db0967"
                        )
            raise ValueError("Invalid values. Expected int")
        self.Name_display.configure(text=self.Name_entry.get())
        self.Date_today_display.configure(text=date.today().strftime("%a-%b-%Y"))
        if self.radio_var.get() == 2:
            self.Gender_display.configure(text="Female")
        elif self.radio_var.get() == 1:
            self.Gender_display.configure(text="Male")
        else:
            pass

        self.Table.insert(column=0, row=0,value="Subjects")
        self.Table.insert(column=1, row=0,value="Scores")

        for i in self.scores["Subjects"]:
            print(i)
            self.Table.insert(column=0, row=(self.scores["Subjects"].index(i)+1),value=i)

        r = 1
        for i in (self.scores["Scores"]):
            self.Table.insert(column=1, row=r,value=i)
            print(f"inserted {i} at {r}" )
            r += 1
        
        CTkMessagebox(master=self.Data_Frame,title="Submitted", icon="info",message="Data submitted",icon_size=(50,50),corner_radius=25,sound=True,font=("century gothic",20, "bold"),header=False,topmost=True,text_color="#db0967")
        self.Table.pack(expand=True, fill="both", side="top")
        self.analysis()

    def analysis(self):
        self.df = pd.DataFrame(self.scores)
        average = round(self.df.mean(numeric_only=True),2).to_string(dtype=False, header=False, index=False, name=False)
        sorted = self.df.sort_values(by="Scores", ascending=False)
        best = sorted.iloc[0]["Subjects"]
        worst = sorted.iloc[-1]["Subjects"]
        self.Average_score_display.configure(text=average)
        self.Best_subject_display.configure(text=best)
        self.worst_subject_display.configure(text=worst)

        # Clear any existing plots in the frames
        for widget in self.graph_1.winfo_children():
            widget.destroy()
        for widget in self.graph_2.winfo_children():
            widget.destroy()
        for widget in self.graph_3.winfo_children():
            widget.destroy()
        for widget in self.graph_4.winfo_children():
            widget.destroy()

        # Plotting and displaying graphs
        font = {'family': 'Arial', 'size': 18, 'weight': 'bold'}
        font2 = {'family': 'Arial', 'size': 12, 'weight': 'bold'}

        self.fig1 = plt.figure(figsize=(9,4), dpi=80)
        self.ax1 = self.fig1.add_subplot(111)
        self.ax1.fill_between(self.df["Subjects"], self.df["Scores"], color="#db0967")
        self.ax1.set_title("Performance", fontdict=font)
        self.ax1.set_xlabel("Subjects", fontdict=font2)
        self.ax1.set_ylabel("Scores", fontdict=font2)

        canvas1 = FigureCanvasTkAgg(self.fig1, master=self.graph_1)
        canvas1.draw()
        canvas1.get_tk_widget().pack(fill="both", expand=True)

        self.fig2 = plt.figure(figsize=(9,4), dpi=80)
        self.ax2 = self.fig2.add_subplot(111)
        self.ax2.plot(self.df["Subjects"], self.df["Scores"], color="#db0967")
        self.ax2.set_title("Performance", fontdict=font)
        self.ax2.set_xlabel("Subjects", fontdict=font2)
        self.ax2.set_ylabel("Scores", fontdict=font2)

        canvas2 = FigureCanvasTkAgg(self.fig2, master=self.graph_2)
        canvas2.draw()
        canvas2.get_tk_widget().pack(fill="both", expand=True)

        self.fig3 = plt.figure(figsize=(9,4), dpi=80)
        self.ax3 = self.fig3.add_subplot(111)
        self.ax3.barh(self.df["Subjects"], self.df["Scores"], color="#db0967")
        self.ax3.set_title("Performance", fontdict=font)
        self.ax3.set_xlabel("Subjects", fontdict=font2)
        self.ax3.set_ylabel("Scores", fontdict=font2)

        canvas3 = FigureCanvasTkAgg(self.fig3, master=self.graph_3)
        canvas3.draw()
        canvas3.get_tk_widget().pack(fill="both", expand=True)

        self.fig4 = plt.figure(figsize=(9,4), dpi=80)
        self.ax4 = self.fig4.add_subplot(111)
        self.ax4.bar(self.df["Subjects"], self.df["Scores"], color="#db0967")
        self.ax4.set_title("Performance", fontdict=font)
        self.ax4.set_xlabel("Subjects", fontdict=font2)
        self.ax4.set_ylabel("Scores", fontdict=font2)

        canvas4 = FigureCanvasTkAgg(self.fig4, master=self.graph_4)
        canvas4.draw()
        canvas4.get_tk_widget().pack(fill="both", expand=True)
        # Set background color of plots
        self.ax1.xaxis.label.set_color("white")
        self.ax1.yaxis.label.set_color("white")
        self.ax1.title.set_color("#db0967")
        self.ax1.tick_params(axis='x', colors="#db0967")
        self.ax1.tick_params(axis='y', colors="#db0967")
        self.ax1.spines['top'].set_color("#2b2b2b")
        self.ax1.spines['right'].set_color("#2b2b2b")
        self.ax1.spines['bottom'].set_color("white")
        self.ax1.spines['left'].set_color("white")

        for tick in self.ax1.get_xticklabels() + self.ax1.get_yticklabels():
            tick.set_fontname('Arial')
            tick.set_fontsize(11)
            tick.set_fontweight('bold')

        self.ax2.xaxis.label.set_color("white")
        self.ax2.yaxis.label.set_color("white")
        self.ax2.title.set_color("#db0967")
        self.ax2.tick_params(axis='x', colors="#db0967")
        self.ax2.tick_params(axis='y', colors="#db0967")
        self.ax2.spines['top'].set_color("#2b2b2b")
        self.ax2.spines['right'].set_color("#2b2b2b")
        self.ax2.spines['bottom'].set_color("white")
        self.ax2.spines['left'].set_color("white")

        for tick in self.ax2.get_xticklabels() + self.ax2.get_yticklabels():
            tick.set_fontname('Arial')
            tick.set_fontsize(11)
            tick.set_fontweight('bold')

        self.ax3.xaxis.label.set_color("white")
        self.ax3.yaxis.label.set_color("white")
        self.ax3.title.set_color("#db0967")
        self.ax3.tick_params(axis='x', colors="#db0967")
        self.ax3.tick_params(axis='y', colors="#db0967")
        self.ax3.spines['top'].set_color("#2b2b2b")
        self.ax3.spines['right'].set_color("#2b2b2b")
        self.ax3.spines['bottom'].set_color("white")
        self.ax3.spines['left'].set_color("white")

        for tick in self.ax3.get_xticklabels() + self.ax3.get_yticklabels():
            tick.set_fontname('Arial')
            tick.set_fontsize(11)
            tick.set_fontweight('bold')

        self.ax4.xaxis.label.set_color("white")
        self.ax4.yaxis.label.set_color("white")
        self.ax4.title.set_color("#db0967")
        self.ax4.tick_params(axis='x', colors="#db0967")
        self.ax4.tick_params(axis='y', colors="#db0967")
        self.ax4.spines['top'].set_color("#2b2b2b")
        self.ax4.spines['right'].set_color("#2b2b2b")
        self.ax4.spines['bottom'].set_color("white")
        self.ax4.spines['left'].set_color("white")

        for tick in self.ax4.get_xticklabels() + self.ax4.get_yticklabels():
            tick.set_fontname('Arial')
            tick.set_fontsize(11)

        for i in [self.fig1,self.fig2,self.fig3,self.fig4,
                self.ax1,self.ax2,self.ax3,self.ax4]:
            i.set_facecolor("#2b2b2b")

    def import_file(self):
        file_pth = filedialog.askopenfilename(defaultextension='.xlsx',filetypes=[("Excel Files", "*.xlsx"),],title='Select an Excel file')
        
        

    def on_closing(self):
        msg = CTkMessagebox(message=" Are you Leaving??", title="Quit?", option_1="Leaving", option_2="Staying", icon="question")
        
        if msg.get()=="Leaving":
            """To safely close graphs on exit inorder to avoid error."""
            try:
                plt.close(fig=self.fig1)
                plt.close(fig=self.fig2)
                plt.close(fig=self.fig3)
                plt.close(fig=self.fig4)
            except:
                print("Graphs not yet plotted.\nExiting")
            finally:
                self.root.quit()

if __name__ == "__main__":
    splash()
    app = ReportApp()
    app.run()
