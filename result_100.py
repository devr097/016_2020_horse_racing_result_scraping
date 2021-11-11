# coding:utf-8

import _tkinter
import os
import platform
import subprocess
import sys
import tkinter
import csv
# from openpyxl import Workbook
import openpyxl
import xlsxwriter
from datetime import datetime
from tkinter import filedialog
from tkinter import messagebox
from shutil import copyfile
from tkinter.scrolledtext import ScrolledText
from tkinter import ttk

class App(tkinter.Tk): 
    
    def __init__(self):
        super(App, self).__init__()
        # init windows
        self.title("Scrap on Racing Result")
        self.geometry("800x800")
        self.minsize(200, 400)
        self.resizable(False, True)
        self.config(background="#f2f2f2")
        
        # //////////////////////////////////// Top frame
        self.topframe = tkinter.LabelFrame(self, text="Informations", height=150)
        self.topframe.pack(pady=30, ipadx=15, ipady=10)
                
        self.url_label = tkinter.Label(self.topframe, text="URL List File")
        self.url_label.pack()        
        self.btn_select = tkinter.Button(self.topframe, text="Select file", width="25",
                                         command=self.select_file).pack(pady=10)
        self.url_fname_var = tkinter.StringVar()
        self.url_fname = tkinter.Entry(self.topframe, width=60,textvariable=self.url_fname_var)
        self.url_fname.pack()
        
       
        # //////////////////////////////////// Output frame
        self.output = tkinter.LabelFrame(self, text="Result", height=200)
        self.output.pack(ipadx=15, ipady=10)
        # widgets
        self.progress = tkinter.DoubleVar()
        self.progress_counter_var = tkinter.IntVar(value=0)
        self.progress_total_var = tkinter.IntVar(value=0)
        self.progress_bar = ttk.Progressbar(self.output, variable=self.progress, length=400)
        self.progress_bar.pack(pady=20)
        self.progress_counter = tkinter.Label(self.output, textvariable=self.progress_counter_var)   
        self.progress_total = tkinter.Label(self.output, textvariable=self.progress_total_var)
       
        # self.progress_bar.start(10)
        self.response_entry = ScrolledText(self.output, width="90", height="26", undo=True)
        self.response_entry.pack(padx=10, expand=tkinter.TRUE, fill="both")
        # btn
        
        self.btn_search = tkinter.Button(self.output, text="Research", width="25",
                                         command=self.call_spider_process).pack(pady=10)
        # self.btn_search = tkinter.Button(self.output, text="Research", width="25",
        #                                  command=self._save_file()).pack(pady=10)
        self.btn_close = tkinter.Button(self.output, text="Close", width="25", command=self.quit).pack(pady=10,side="bottom")
        self.btn_save_actived = True

    def call_spider_process(self):
        
        """
        Call the subprocess command fo call scrapy
        :return:
        """
        
        try:
            self.response_entry.delete(1.0, "end")  # reset response output
            self.response_entry.update()
            self.progress_counter_var.set(0)
            self.progress_total_var.set(0)
            self.progress_bar.step(0)
            self.progress_bar.update()
            self.progress.set(0)
            if os.path.isfile("./tmp/130.csv"):
                 os.remove("./tmp/130.csv")
        except _tkinter.TclError:
            raise
                    

        if os.path.isfile(self.url_fname.get()):
            urllist = []
            dstFilename = self.url_fname.get()
           
            with open(dstFilename) as csvfile:
                readCSV = csv.reader(csvfile)
                url_number = 0
                for row in readCSV:           
                    if len(row)==1:
                        if row[0][0:4] == 'http': 
                            urllist.append(row)
                            url_number = url_number + 1
                        else:
                            pass
                    else:
                        pass
                if url_number == int(chr(48)):
                    tkinter.messagebox.showerror("Error", "There is no correct url in the current file.\nPlease select url list file again.")
                    return
                else:
                    pass     
            if url_number < int(chr(49)+chr(48)+chr(49)):       
                cur_no = 0
                for turl in urllist:
                    cur_url = turl[0].strip()
                    cur_no = cur_no + 1
                    if os.path.isfile("./tmp/resultcsv.csv"):
                        os.remove("./tmp/resultcsv.csv")
                    if cur_no ==1:
                        self.response_entry.insert(tkinter.END, "\nStarting...\n")
                        
                    command = "scrapy", "runspider", "tkspider.py", "-a", "url=" + cur_url

                    subprocess.run(command) 
                    
                    self.display_result(cur_url, url_number, cur_no)
                    self.SaveResult("./tmp/resultcsv.csv", "./tmp/130.csv")

                self.create_save_btn()
           
        else:
            tkinter.messagebox.showerror("Error", "This file doesn't exist.\nPlease select url list file again. ")

    def SaveResult(self, srcfile, dstfile):

         srclist = []
         with open(srcfile) as csvfile:
                readCSV = csv.reader(csvfile)
                for row in readCSV:
                    srclist.append(row)

         with open(dstfile,'a') as resultfile:
                writer = csv.writer(resultfile)                
                for i in srclist:
                    writer.writerow(i)
                
    def display_result(self, cur_url, total_no, cur_no):                
       
        file = "./tmp/resultcsv.csv"

       # check if file exist and if not empty
        if os.path.isfile(file) and os.stat(file).st_size != 0:
            
            string_revised = cur_url.ljust(85)
            tmp_str = "  " + string_revised + "End!\n"

            self.response_entry.insert(tkinter.END, tmp_str)
            
            if cur_no == total_no:
                self.response_entry.insert(tkinter.END, "\nCompleted!")
            
            self.progress_total.place(x=550, y=17)
            self.progress_total_var.set(total_no)

            self.progress_counter_var.set(cur_no)  
            self.progress_counter.place(x=350, y=10)
            
            self.progress_bar.step(100 / total_no)
            self.progress_bar.update()
            self.progress.set(cur_no/total_no*100)            
        
    def create_save_btn(self):
        """
        Create Save btn
        :return:
        """
        if self.btn_save_actived:
            btn_save = tkinter.Button(self.output, text="Save", width="25", bg="#8ce261",
                                      command=self._save_file)
            btn_save.pack(side="top")
        self.btn_save_actived = False
    
    
    def convert_csv_to_xlsx(self,csv_file, xls_file):
        workbook=xlsxwriter.Workbook(xls_file)
        worksheet=workbook.add_worksheet()
        category=["TYPE OF RACE",'DATE','COURSE','DISTANCE/Y','GOING','CLASS','POSITION','HORSE NAME','AGE','WEIGHT','STONES','POUNDS','ALL POUNDS','JCK ALNC','3YO ALLOWANCE','WON','SCORE','COMMENTS','MARGIN']
        date_format=workbook.add_format({'num_format':'dd/mm/Y','align':'right'})
        color01_format=workbook.add_format({'bold':True,'border':3,'align':'left','fg_color':'#ccf2ff'})
        color02_format=workbook.add_format({'bold':True,'border':3,'fg_color':'#ffff99'})
        cal01_format=workbook.add_format({'bold':True,'border':3,'fg_color':'#ccf2ff','font_color':'#ff0000'})
        cal02_format=workbook.add_format({'bold':True,'border':3,'bg_color':'#00ccff'})
        cal03_format=workbook.add_format({'bold':True,'border':3,'bg_color':'#ff9900'})
        cal04_format=workbook.add_format({'bold':True,'border':3,'bg_color':'#009900'})

        with open(csv_file) as f:
            reader = csv.reader(f)
            spaceline = 0
            score_row=0
            for r, row in enumerate(reader):
                
                for c, col in enumerate(row):                    
                    if "https" in str(col):     
                        if r==0:
                            pass
                        else:
                            spaceline=spaceline+2  
                        worksheet.write(r+spaceline+1, c+0,col)
                        worksheet.write(r+spaceline+1, c+16, None)                        
                        worksheet.write(r+spaceline+0, c+16, None)
                        worksheet.write(r+spaceline+2, c+16, None)
                        worksheet.write(r+spaceline+1, c+12, None)
                        worksheet.write(r+spaceline+0, c+12, None)
                        worksheet.write(r+spaceline+2, c+12, None)
                    elif col=="TYPE OF RACE":
                        typeafter=1
                        c_number=0
                        for item in category:
                            if c_number>8 and c_number<13:
                                worksheet.write(r+spaceline+2, c_number, item, color02_format)     
                            else:                  
                                worksheet.write(r+spaceline+2, c_number, item, color01_format)    
                            c_number += 1                                                       
                        worksheet.write(r+spaceline+2,c_number,'Ref O.R.',cal01_format)
                        worksheet.write(r+spaceline+2,c_number+1,'Ref Pnds',cal01_format)
                        worksheet.write(r+spaceline+2,c_number+2,'Ref Scr',cal01_format)
                        worksheet.write_formula(r+spaceline+3,c_number,chr(57)+chr(48),cal01_format)
                        worksheet.write_formula(r+spaceline+3,c_number+1,chr(49)+chr(52)+chr(53),cal01_format)
                        worksheet.write_formula(r+spaceline+3,c_number+2,chr(53)+chr(48),cal01_format)
                        worksheet.write(r+spaceline+4,c_number,'O.R.')
                        worksheet.write(r+spaceline+4,c_number+1,'All Pounds')
                        worksheet.write(r+spaceline+4,c_number+2,'Ref Scr')
                        worksheet.write_formula(r+spaceline+5,c_number,chr(45)+chr(49)+chr(48)+chr(48),cal02_format)                        
                        equationpound=chr(77)+str(r+spaceline+3+1)+'+'+chr(78)+str(r+spaceline+3+1)+'+'+chr(79)+str(r+spaceline+3+1)
                        print(equationpound,equationpound)
                        worksheet.write_formula(r+spaceline+5,c_number+1,equationpound,cal03_format)
                        equation06=chr(61)+chr(86)+str(r+spaceline+4)+'+'+chr(40)+chr(84)+str(r+spaceline+4)+chr(45)+chr(84)+str(r+spaceline+6)+chr(41)+chr(45)+chr(40)+chr(85)+str(r+spaceline+4)+chr(45)+chr(85)+str(r+spaceline+6)+')'
                        worksheet.write_formula(r+spaceline+5,c_number+2,equation06,cal04_format)
                        
                        score_row = r+spaceline+5
                        break                     
                    else:
                        if c==3 or c==6 or c==8 or c == 10 or c==11 or c==12 or c == 16 or c==18:
                            if c==12:
                                equation='=K'+str(r+spaceline+3)+'*14+'+'L'+str(r+spaceline+3)
                                worksheet.write_formula(r+spaceline+2,c+0,equation)
                            elif c==16 and typeafter==1:
                                typeafter = 0
                                equation='=V'+str(score_row+1)
                                worksheet.write_formula(r+spaceline+2,c+0, equation)                                                   
                            elif c==16:
                                equationscore='=Q'+str(score_row-1)+'+'+'S'+str(r+spaceline+3)+'*4'
                                worksheet.write_formula(r+spaceline+2,c+0,equationscore)
                            elif c==17 and r==2:
                                equationcomments='T'+str(score_row+1)
                                worksheet.write_formula(r+spaceline+2,c+0,equationcomments)
                            else:
                                worksheet.write(r+spaceline+2, c+0, float(col))                                                
                        else:
                            if c==13 or c==14 or c==15:
                                if col!="":
                                    worksheet.write(r+spaceline+2, c+0, float(col)) 
                            elif c==1:
                                col=col.split('/')
                                date=datetime.strptime(col[2]+'-'+col[1]+'-'+col[0],'%Y-%m-%d')
                                date_time = datetime.strptime('2013-01-23 12:30:05.123',
                                                            '%Y-%m-%d %H:%M:%S.%f')
                                worksheet.write_datetime(r+spaceline+2, c,date, date_format)
                            else:
                                worksheet.write(r+spaceline+2, c+0, col)
            # sheet.delete_rows(r+4+spaceline,2570-r-4-spaceline)
        workbook.close()
        
    def _save_file(self):
        """
        Save file and copy the content of tkspider files
        :return:
        """
        file = filedialog.asksaveasfile(mode="w", initialdir="./",
                                    defaultextension=f".xlsx",
                                    filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
        
        spider_file = "./tmp/130.csv"
        self.convert_csv_to_xlsx(spider_file, file.name)
      
    def select_file(self):
        """
        Select file include 10 URLS
        :return:
        """
        
        file = filedialog.askopenfilename(initialdir="./", title = "Select File",
                                    filetypes=(("Excel files", "*.csv"), ("All files", "*.*")))
        
        self.url_fname_var.set(file)
      
   
if __name__ == "__main__":
    app = App()
    app.mainloop()
