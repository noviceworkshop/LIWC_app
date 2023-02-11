import os
import re
import glob
import time
import codecs
from tkinter import *
import pandas as pd
from tkinter import filedialog
import tkinter.messagebox
from openpyxl import load_workbook
import string
import platform
from collections import Counter
is_windows = platform.system() == 'Windows'

from ckiptagger import WS, POS

def replacing(candidates, text):
    for candidate in candidates:
        text = text.replace(*candidate)
    return text

def get_titles(limit=1000):
    i = 1
    di = {}
    for i1 in string.ascii_uppercase:
        di[i1] = i
        if i == limit:
            return di
        i += 1
    for i1 in string.ascii_uppercase:
        for i2 in string.ascii_uppercase:
            di[i1 + i2] = i
            if i == limit:
                return di
            i += 1
    for i1 in string.ascii_uppercase:
        for i2 in string.ascii_uppercase:
            for i3 in string.ascii_uppercase:
                di[i1 + i2 + i3] = i
                if i == limit:
                    return di
                i += 1
    


class GUIDemo(Frame):
    def __init__(self, master=None,Freq=None):
        Frame.__init__(self, master)
        self.grid()
        self.createWidgets()
        self.fin = None
        self.Freq = Freq
        self.gex = re.compile(r"\((\w+)\)")
        self.replacements = [['“', '"'],['”', '"'],['。', '.'], ['，', ','], ['、', ','], ['；', ';'], ['：', ':'], ['「', '"'], ['」', '"'], ['『', '"'], ['』', '"'], ['（', '('], ['）', ')'], ['【', '('], ['】', ')'], ['？', '?'], ['！', '!'], ['──', '-'], ['─', '-'], ['……', '…'], ['—', '-'], ['＃', '#'], ['＄', '$'], ['％', '%'], ['＆', '&'], ['＊', '*'], ['＋', '+'], ['／', '/'], ['＜', '<'], ['＝', '='], ['＞', '>'], ['^', '^'], ['﹍', '_'], ['～', '~'], ['\u3000', ' ']]
        self.pos_list = [['A', 'c_A'], ['Caa', 'c_Caa'], ['Cab', 'c_Cab'], ['Cba', 'c_Cba'], ['Cbb', 'c_Cbb'], ['Da', 'c_Da'], ['Dfa', 'c_Dfa'], ['Dfb', 'c_Dfb'], ['Di', 'c_Di'], ['Dk', 'c_Dk'], ['D', 'c_D'], ['Na', 'c_Na'], ['Nb', 'c_Nb'], ['Nc', 'c_Nc'], ['Ncd', 'c_Ncd'], ['Nd', 'c_Nd'], ['Neu', 'c_Neu'], ['Nes', 'c_Nes'], ['Nep', 'c_Nep'], ['Neqa', 'c_Neqa'], ['Neqb', 'c_Neqb'], ['Nf', 'c_Nf'], ['Ng', 'c_Ng'], ['Nh', 'c_Nh'], ['Nv', 'c_Nv'], ['I', 'c_I'], ['P', 'c_P'], ['T', 'c_T'], ['VA', 'c_VA'], ['VAC', 'c_VAC'], ['VB', 'c_VB'], ['VC', 'c_VC'], ['VCL', 'c_VCL'], ['VD', 'c_VD'], ['VE', 'c_VE'], ['VF', 'c_VF'], ['VG', 'c_VG'], ['VH', 'c_VH'], ['VHC', 'c_VHC'], ['VI', 'c_VI'], ['VJ', 'c_VJ'], ['VK', 'c_VK'], ['VL', 'c_VL'], ['V_2', 'c_V_2'], ['DE', 'c_DE'], ['SHI', 'c_SHI'], ['FW', 'c_FW'], ['COMMACATEGORY', 'c_comma'], ['DASHCATEGORY', 'c_dash'], ['ETCCATEGORY', 'c_etc'], ['EXCLAMATIONCATEGORY', 'c_exclam'], ['PARENTHESISCATEGORY', 'c_parenth'], ['PAUSECATEGORY', 'c_pause'], ['PERIODCATEGORY', 'c_period'], ['QUESTIONCATEGORY', 'c_qmark'], ['COLONCATEGORY', 'c_colon'], ['SEMICOLONCATEGORY', 'c_semic'], ['SPCHANGECATEGORY', 'c_sp']]
        self.comma_dict = {'COMMACATEGORY': 'c_comma', 'DASHCATEGORY': 'c_dash', 'ETCCATEGORY': 'c_etc', 'EXCLAMATIONCATEGORY': 'c_exclam', 'PARENTHESISCATEGORY': 'c_parenth', 'PAUSECATEGORY': 'c_pause', 'PERIODCATEGORY': 'c_period', 'QUESTIONCATEGORY': 'c_qmark', 'COLONCATEGORY': 'c_colon', 'SEMICOLONCATEGORY': 'c_semic', 'SPCHANGECATEGORY': 'c_sp'}
        self.column_titles = get_titles(1000)
        self.ws_root = None
        self.ws = None
        self.pos = None

    def set_ckip_root(self):
        path = filedialog.askdirectory(initialdir=os.getcwd())
        if len(path) == 0:
            return
        self.ws_root = path
        self.ws = path
        self.CkipText["text"] = "--- CkipTagger 初始化中 ---"
        try:
            self.ws = WS(path)
            try:
                self.pos = POS(path)
            except:
                "Error: POS失敗，再按一次"
            # path = path.replace("/", "\\")
            self.CkipText["text"] = "--- 設定CKIP資料夾為 ---"
            self.CkipText2["text"] = f"路徑:{path}"
            tkinter.messagebox.showinfo("CKIP設定",  "initialize 成功!")
        except:
            self.CkipText["text"] = "Error: 初始化CKIP資料夾失敗"



    def detect_coding(self, filename):
        candidates = ['utf-8', 'cp950', 'big5-hkscs', 'utf-16']
        for candidate in candidates:
            try:
                f = codecs.open(filename, encoding=candidate)
                _ = f.readline()
                return candidate
            except:
                continue
        self.displayText2["text"] = f"File `{os.path.split(filename)[1]}` encoding not by UTF-8 or ANSI."
        return None

    # def choose(self):
    #     self.fin = filedialog.askopenfilename()
    #     if self.fin == "":
    #         return
    #     basename, filename = os.path.split(self.fin)
    #     outname = os.path.join("result", "fmt_"+filename)
    #     self.displayText["text"] = "Input file: "+filename
    #     #self.displayText2["text"] = "Output to: "+outname

    def choose_files(self):
        self.fin = filedialog.askopenfilenames(filetypes=[("Text Files", "*.xlsx"),("Text Files", "*.xls")])
        if len(self.fin) == 0:
            return
        filenames = [os.path.split(fin)[1] for fin in self.fin]
        #outnames = [os.path.join("result", "fmt_"+filename) for filename in filenames]
        self.displayText["text"] = "---- Input File ----\n"+"\n".join(filenames)
        self.displayText2["text"] = ""

    def choose_dir(self):
        mydir = filedialog.askdirectory(initialdir=os.getcwd())
        if len(mydir) == 0:
            return
        self.fin = glob.glob(os.path.join(mydir, "*.xlsx")) + glob.glob(os.path.join(mydir, "*.xls"))
        filenames = [os.path.split(fin)[1] for fin in self.fin]
        self.displayText["text"] = "---- Input File ----\n" + "\n".join(filenames)
        self.displayText2["text"] = ""

    
    def choose_files_txt(self):
        self.fin = filedialog.askopenfilenames(filetypes=[("Text Files", "*.txt")])
        if len(self.fin) == 0:
            return
        filenames = [os.path.split(fin)[1] for fin in self.fin]
        #outnames = [os.path.join("result", "fmt_"+filename) for filename in filenames]
        self.displayText["text"] = "---- Input File ----\n"+"\n".join(filenames)
        self.displayText2["text"] = ""

    def choose_dir_txt(self):
        mydir = filedialog.askdirectory(initialdir=os.getcwd())
        if len(mydir) == 0:
            return
        self.fin = glob.glob(os.path.join(mydir, "*.txt"))
        filenames = [os.path.split(fin)[1] for fin in self.fin]
        #outnames = [os.path.join("result", "fmt_"+filename) for filename in filenames]
        self.displayText["text"] = "---- Input File ----\n" + "\n".join(filenames)
        self.displayText2["text"] = ""

    def assert_msg(self, cond, text):
        if cond:
            return True
        else:
            self.displayText2["text"] = text
            return False

    def check_fileType(self):
        start = time.time()
        if type(self.fin) in [list,tuple]:
            files = self.fin
        elif type(self.fin) == str:
            files = [self.fin]
        else:
            self.displayText2["text"] = "Unknown filetype " + str(type(self.fin)) + ": " + str(self.fin)
            return
        # done_seg = False
        # if self.ws is None and not done_seg:
        if self.ws is None:
            if self.ws_root is None:
                self.displayText2["text"] = "Error: CKIP 資料夾未設定"
            else:
                self.displayText2["text"] = "Error: CKIP 初始化未完成或失敗"
            return
        if files[0][-5:] == '.xlsx' or files[0][-4:] == '.xls':
            donefile = self.process_excel(files)
        elif files[0][-4:] == '.txt':
            donefile = self.process_txt(files)
        if donefile !=False: #did not have error during the process
            period = time.time() - start
            self.displayText2["text"] = "處理完成 (in %.3f s)"%period
            self.displayText["text"] = "---- Output File ----\n" + '\n'.join(donefile)
            self.displayText2["text"] = "處理完成 (in %.3f s)"%period
    def process_excel(self,file_list):
        
        donefile = []
        column = self.column_entry.get().upper()
        if not self.assert_msg(column in self.column_titles, "Unknown column " + column):
            return False
        column_idx = self.column_titles[column]-1
        column_name = ['c_filename']
        column_name.extend(dict(self.pos_list).values())
        df = pd.DataFrame(columns=column_name)
        for idx, filename in enumerate(file_list):
            df = pd.read_excel(filename) #xls uses xlrd, xlsx uses openpyxl
            basename, filename = os.path.split(filename)
            freq_df = df.copy()
            if not self.assert_msg(column_idx <= df.shape[1], "column index out of range"):
                    return False
            tmp_txts = []
            tmp_rows = []
            for row in range(0, df.shape[0]):
                text = df.iloc[row,column_idx]
                if type(text)!=str:
                    text = str(text)
                text = text.replace('\r\n', '')
                text = text.replace('\n', '')
                tmp_txts.append(text)
                tmp_rows.append(row)
            seg_txts = self.ws(tmp_txts)
            print('finish seg: ',len(seg_txts))
            for row, text in zip(tmp_rows, seg_txts):
                text = ' '.join(text)
                text = replacing(self.replacements, text)
                df.loc[row,f'seg_{column}']=text    
            os.makedirs(os.path.join(basename, 'result'), exist_ok=True)
            if filename[-4:] == '.xls':
                filename =filename[:-4]+'_xls.xlsx' #make the save file xlsx
            filename = f"fmt_seg_{column}_{filename}"
            outname = os.path.join(basename, 'result', filename)
            print(outname)
            df.to_excel(outname,index=False)
            print('seg exported')
            del df

            #####frequency
            if self.Freq==True:
                pos_txts = self.pos(seg_txts)
                print('finish pos: ' , len(pos_txts))

                freq_df[list(dict(self.pos_list).values())]=0
                for row, pos in zip(tmp_rows, pos_txts):
                    pos= [dict(self.pos_list)[i] for i in pos if i not in ['WHITESPACE','DOTCATEGORY']]
                    c = Counter(pos)
                    freq = dict([(i, c[i] / len(pos) * 100.0) for i in c])
                    for k,v in freq.items():
                        freq_df.loc[row,k]=v
                freq_df.to_excel(os.path.join(basename,'result',filename[:-5]+'_frequency.xlsx'),index=False)
                print('pos exported')
            donefile.append(os.path.join('result', 'fmt_'+filename))
        return donefile


    def process_txt(self,file_list):
        donefile = []
        for idx, filename in enumerate(file_list):
            codec = self.detect_coding(filename)
            if codec is None:
                return
            fin = open(filename, 'r', encoding=codec)
            basename, filename = os.path.split(filename)
            os.makedirs(os.path.join(basename, 'result'), exist_ok=True)
            outname = os.path.join(basename, 'result', 'fmt_'+filename)
            fout = open(outname, 'w', encoding='utf-8-sig', newline='\n')
            tmp_txts = []
            for line in fin:
                line = line.replace('\r\n', '')
                line = line.replace('\n', '')
                tmp_txts.append(line)
            seg_txts = self.ws(tmp_txts)
            
            for line in seg_txts:
                line = ' '.join(line)
                line = replacing(self.replacements, line)
                line = re.sub(r"(: | : )",' : ',line)
                fout.write(line.strip()+'\n')
            donefile.append(os.path.join('result', 'fmt_'+filename))
            if self.Freq==True:
                pos_txt = self.pos(seg_txts)[0]
                pos_txt = [dict(self.pos_list)[i] for i in pos_txt if i not in ['WHITESPACE','DOTCATEGORY']]
                c = Counter(pos_txt)
                freq=dict(zip(dict(self.pos_list).values(),[0]*len(self.pos_list)))        
                for i in c:
                    freq[i]=c[i]/len(pos_txt)* 100.0
                    # print(c[i],i,freq[i])
                    freq['c_filename'] = filename
                df = df.append(freq,ignore_index=True)
                df.to_excel(os.path.join(basename,'result','txt_frequency.xlsx'),index=False)
        return donefile

                
        
    def createWidgets(self):
        row = 1
        self.topText = Label(self, width=20, height=1)
        self.topText["text"] = "\u3000"
        self.topText.grid(row=row, column=0, columnspan=5)
        row += 1
        
        self.gap = Label(self, width=5, height=1)
        self.gap.grid(row=row, column=0)

        self.ckip_button = Button(self, width=20, height=1)
        self.ckip_button["text"] = "設定CKIP資料夾"
        self.ckip_button.grid(row=row, column=1, columnspan=2)
        self.ckip_button["command"] = self.set_ckip_root
        row += 1

        self.CkipText = Label(self)
        self.CkipText["text"] = "--- CKIP資料夾 未設定 ---"
        self.CkipText.grid(row=row, column=0, columnspan=7)
        row += 1
        self.CkipText2 = Label(self)
        self.CkipText2["text"] = ""
        self.CkipText2.grid(row=row, column=0, columnspan=7)
        row += 1

        self.HeadText2 = Label(self)
        self.HeadText2["text"] = "__________ 純文字文件 __________"
        self.HeadText2.grid(row=row, column=0, columnspan=7)
        row += 1

        self.register_button3 = Button(self, width=9, height=1)
        self.register_button3["text"] = "開啟文字檔"
        self.register_button3.grid(row=row, column=1)
        self.register_button3["command"] = self.choose_files_txt


        self.register_button4 = Button(self, width=9, height=1)
        self.register_button4["text"] = "開啟資料夾"
        self.register_button4.grid(row=row, column=2)
        self.register_button4["command"] = self.choose_dir_txt

        self.gap3 = Label(self, width=5, height=1)
        self.gap3.grid(row=row, column=3)
        row += 1


        self.HeadText = Label(self)
        self.HeadText["text"] = "__________ EXCEL文件 __________"
        self.HeadText.grid(row=row, column=0, columnspan=7)
        row += 1

        self.register_button1 = Button(self, width=9, height=1)
        self.register_button1["text"] = "開啟xlsx檔"
        self.register_button1.grid(row=row, column=1)
        self.register_button1["command"] = self.choose_files


        self.register_button2 = Button(self, width=9, height=1)
        self.register_button2["text"] = "開啟資料夾"
        self.register_button2.grid(row=row, column=2)
        self.register_button2["command"] = self.choose_dir

        self.gap2 = Label(self, width=5, height=1)
        self.gap2.grid(row=row, column=3)
        row += 1

        self.column_label = Label(self, text='輸入欲處理欄位')
        self.column_label.grid(row=row, column=1)
        self.column_variable = StringVar(self, value='B')
        self.column_entry = Entry(self, width=10, textvariable=self.column_variable)
        self.column_entry.grid(row=row, column=2)
        row += 1
        

        # self.done_label = Label(self, text='是否已斷詞過')
        # self.done_label.grid(row=row, column=1)
        # self.done_variable = StringVar(self, value='No')
        # self.done_segmt = OptionMenu(self, self.done_variable, "Yes", "No")
        # self.done_segmt.grid(row=row, column=2)
        # row += 1
                
        self.displayText3 = Label(self)
        self.displayText3["text"] = ""
        self.displayText3.grid(row=row, column=0, columnspan=7)
        row += 1
        
        self.unlock_button = Button(self, width=20, height=1)
        self.unlock_button["text"] = "開始處理"
        self.unlock_button.grid(row=row, column=1, columnspan=2)
        self.unlock_button["command"] = self.check_fileType
        row += 1
        
        self.displayText = Label(self)
        self.displayText["text"] = ""
        self.displayText.grid(row=row, column=0, columnspan=7)
        row += 1

        self.displayText2 = Label(self)
        self.displayText2["text"] = ""
        self.displayText2.grid(row=row, column=0, columnspan=7)
        row += 1
