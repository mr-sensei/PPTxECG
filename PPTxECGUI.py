import os, time, logging
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import PPTxECG as pptxecg
#logging.basicConfig(level=logging.DEBUG, format=' %(levelname)s - %(asctime)s: %(message)s')


###Some strings used in display.
ANALYSIS_PLACEHOLDER_TEXT = """No data.

Please return to the previous tab, select a presentation or parent folder
containing presentations and click 'Analyse'.


"""
SINGLE_PRESENTATION_ANALYSIS = """Presentation: %s
Word count: %s
There are %s slides - Slide #%s is the longest with %s words.
The average slide has %s words.

For more detail, use the button below to create a spreadsheet.
"""
MULTI_PRESENTATION_ANALYSIS = """%s presentations found in %s.
Total number of slides: %s
Total word count: %s
Most verbose presentation:
%s (%s words, %s slides)
For more detail, use the button below to create a spreadsheet.
"""

class GUI():
    def __init__(self):
        '''create instance'''
        self.win = tk.Tk()
        self.win.title("PPT-ECG")
        #Tab1 Vars
        self.root_dir = os.path.dirname(__file__)
        self.file_loc = tk.StringVar()
        self.file_loc.set("<Select a file>")
        self.course_days = tk.IntVar()
        self.course_days.set(5)
        self.course_hours = tk.IntVar()
        self.course_hours.set(7)
        self.hours_calc = tk.StringVar()
        self.hours_calc.set("35")
        self.my_pres = ''
        self.anl_status = tk.StringVar()
        self.anl_status.set("")
        self.data = ()
        #Tab2 Vars
        self.slide_analysis_str = tk.StringVar()
        self.slide_analysis_str.set('')
        self.sv_spreadsheet_status = tk.StringVar()
        self.sv_spreadsheet_status.set("")
        #Initiate
        self.create_outline()
        self.file_entry.focus()

        
    def create_outline(self):
        self.add_menu()
        self.add_tabs()
        self.kthxbai = ttk.Button(self.win,text="Exit",command=self._quit)
        self.kthxbai.grid(row=1,column=0)

    
    def add_tabs(self):
        '''Tab Switch Control'''
        self.tab_switcher = ttk.Notebook(self.win)
        self.file_chooser = ttk.Frame(self.tab_switcher)
        self.file_analysis = ttk.Frame(self.tab_switcher)
        self.tab_switcher.add(self.file_chooser,text="Select Presentation(s)")
        self.tab_switcher.add(self.file_analysis,text="Presentation Analysis")
        #Tab contents called from other methods
        self.chooser_tab(self.file_chooser)
        self.analyser_tab(self.file_analysis)
        self.tab_switcher.grid(row=0, column=0, sticky='NESW')


    def add_menu(self):
        '''Creates a Menu Bar.'''
        self.menu_bar = tk.Menu(self.win)
        self.win.config(menu=self.menu_bar)
        #File menu
        self.file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.file_menu.add_command(label="Check Presentation")
        self.file_menu.add_command(label="About",command=self._about)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Exit", command=self._quit)
        self.menu_bar.add_cascade(label="File", menu=self.file_menu)
    
    #### Individual Tab Content
    def chooser_tab(self, my_tab):
        '''Allows user to specify presentation for analysis.'''
        #Set parameters
        self.course_params = ttk.LabelFrame(my_tab, text=' Set Course Parameters ')
        self.course_params.grid(row=0, column=0, padx=8, pady=4, sticky='NESW')
        #For setting course duration parameters
        self.days_prompt = ttk.Label(self.course_params, text="Course days: ")
        self.days_prompt.grid(row=0, column=0, sticky='W', padx=(10,2), pady=4)
        self.days_entry = ttk.Entry(self.course_params, width=3, textvariable=self.course_days)
        self.days_entry.grid(row=0, column=1, sticky='W', padx=(2,10), pady=4)
        self.hours_prompt = ttk.Label(self.course_params, text="Hours per day: ")
        self.hours_prompt.grid(row=0, column=2, sticky='W', padx=(10,2), pady=4)
        self.hours_entry = ttk.Entry(self.course_params, width=3, textvariable=self.course_hours)
        self.hours_entry.grid(row=0, column=3, sticky='W', padx=(2,10), pady=4)
        self.days_entry.bind("<KeyRelease>",self._calc_hours)
        self.hours_entry.bind("<KeyRelease>",self._calc_hours)
        self.calc_prompt = ttk.Label(self.course_params, text="Course Hours: " + self.hours_calc.get())
        self.calc_prompt.grid(row=0, column=4, sticky='W', padx=(10,2), pady=4)
        #For setting output file parameters
        #Select file
        self.file_picker = ttk.LabelFrame(my_tab, text=' Select Presentation ')
        self.file_picker.grid(row=1, column=0, padx=8, pady=4)
        #For selecting file
        self.file_entry = ttk.Entry(self.file_picker, width=50, textvariable=self.file_loc)
        self.file_entry.grid(row=2, column=0, columnspan = 4, sticky='W', padx=(10), pady=4)
        self.fp_prompt = ttk.Button(self.file_picker, text="Choose file...", command=self._get_file_path)
        self.fp_prompt.grid(row=3, column=0, sticky='W', padx = (10,0), pady=4)
        self.dirpath_prompt = ttk.Button(self.file_picker, text="Choose folder...", command=self._get_folder_path)
        self.dirpath_prompt.grid(row=3, column=1, sticky='W', padx = 0, pady=4)
        ttk.Label(self.file_picker).grid(row=3, column = 2, padx = 30) #padding between choice and analyse button
        self.file_entry.select_range(0, tk.END)
        self.file_entry.bind("<FocusIn>", self._highlight_helper)
        #For analysing file
        self.btn_file_analyse = ttk.Button(self.file_picker, text="Analyse", command=self._analyse_file)
        self.btn_file_analyse.grid(row=3, column=3, sticky = 'E', padx=(0,10))
        self.anl_result = ttk.Label(self.file_picker,textvariable=self.anl_status)
        self.anl_result.grid(row=6, column=0, columnspan=4)
        #Pad out widgets
        self._pad_me(my_tab)

    def analyser_tab(self, my_tab):
        '''Provides some basic output of presentation analysis.'''
        #Frame
        self.slides_output = ttk.LabelFrame(my_tab, text=' Presentation Details ')
        self.slides_output.grid(row=0, column=0, padx=8, pady=4, sticky ='EW')
        #Basic Data
        self.slide_analysis_lbl = ttk.Label(self.slides_output, textvariable=self.slide_analysis_str)
        self.slide_analysis_lbl.grid(row=0, column=0, padx = 4, pady = 4, sticky="EW")
        #Out of frame
        self.btn_make_spreadsheet = ttk.Button(my_tab, text=' Create Summary Spreadsheet ', command=self._request_spreadsheet)
        self.btn_make_spreadsheet.grid(row=1, column=0, sticky='W', padx=(8,0))
        #self.btn_make_spreadsheet.state(["disabled"])
        self.lbl_spreadsheet_status = ttk.Label(my_tab,textvariable=self.sv_spreadsheet_status)
        self.lbl_spreadsheet_status.grid(row=2, sticky='W', padx=(8,0))
        ttk.Label(my_tab, width = 52).grid(row=3) #clumsy way of forcing the output window above to stay bigger than its content.
        self._pad_me(my_tab)
        
        #Engage
        self._update_analysis()


    #### Callbacks and tools
    def _about(self):
        messagebox.showinfo("About",
            """An ECG for PowerPoint. Intended to help bring Powerpoints to life and avoid attentive flatlining caused Death by PowerPoint.""")

    def _analyse_file(self):
        '''Passes file to pptecg.py for analysis'''
        for_analysis = self.file_loc.get()
        logging.debug("Analysing " + for_analysis + "...")
        self.my_pres = pptxecg.analyse_this(for_analysis)
        status = "Unable to analyse."
        if self.my_pres[0] == 0: #if there's no data ditch here
            self.my_pres = ''
            status = "No presentation found. Is this a PPTX file?"
        elif self._unpack_data():
            if self.my_pres[0] > 1:
                status = str(self.my_pres[0]) + " presentations detected and analysed."
            else:
                status = "Presentation analysed."
        _delay_type(self.anl_status, self.anl_result, status)
        #print(self.my_pres)
        return

    
    def _calc_hours(self,key):
        '''Calculates the training hours - disables analyse button if valid number not present'''
        try:
            self.hours_calc.set(str(self.course_days.get() * self.course_hours.get()))
            self.btn_file_analyse.state(["!disabled"])
        except:
            self.hours_calc.set("...")
            self.btn_file_analyse.state(["disabled"])
        self.calc_prompt.configure(text="Course Hours: " + self.hours_calc.get())
    
    def _request_spreadsheet(self):
        logging.debug("Preparing to make spreadsheet.")
        if self.my_pres:
            logging.debug("Looks like it was already analysed: passing in returned metrcis.")
            prez = self.my_pres
        else:
            logging.debug("Doesn't look like it has been analysed yet: passing in file path.")
            prez = self.file_loc.get()
        
        stat = pptxecg.make_spreadsheet_of_this(prez, self.hours_calc.get())
        
        if stat:
            ret = "Spreadsheet created at:\n" + str(stat)
        else:
            ret = "Spreadsheet failed."
        _delay_type(self.sv_spreadsheet_status, self.lbl_spreadsheet_status, ret)
        
    def _get_file_path(self):
        '''gets the path to a file and puts it in the path entry box'''
        self.anl_status.set("")
        fName = filedialog.askopenfilename(parent=self.win, initialdir=self.root_dir, filetypes=[("Powerpoint Open XML","*.pptx"),("All files","*.*")])
        self.file_loc.set(fName)
        #select and highlight
        self.file_entry.select_clear()
        self.btn_file_analyse.focus()
        logging.debug("User selected file: " + self.file_loc.get())
    
    def _get_folder_path(self):
        '''gets the path to a file and puts it in the path entry box'''
        self.anl_status.set("")
        fName = filedialog.askdirectory(parent=self.win, initialdir=self.root_dir)
        self.file_loc.set(fName)
        #select and highlight
        self.file_entry.select_clear()
        self.btn_file_analyse.focus()
        logging.debug("User selected folder: " + self.file_loc.get())
    
    def _highlight_helper(self,event):
        self.file_entry.select_range(0, tk.END)
        self.anl_status.set("")
        
    def _pad_me(self,root):
        for child in root.winfo_children():
            child.grid_configure(padx=8, pady=8)

    def _quit(self):
        '''Function to cleanly exit a given window'''
        self.win.quit()
        self.win.destroy()
        exit()

    def _unpack_data(self):
        '''Unpacks tuple of data into vars.'''
        logging.debug("Extracting returned data...")
        string = ''
        pres_count = self.my_pres[0]
        data = self.my_pres[1]
        if pres_count == 0:
            logging.warning("Somehow a null tuple reached _unpack_data(). Shouldn't happen.")
        elif pres_count == 1:
            logging.debug("_unpack_data() called to unpack a single presentation for analysis tab.")
            for presentation in data:
                word_count = data[presentation][0]
                slide_data = data[presentation][1]
                slide_count = len(slide_data)
                max_slide_wc = 0
                long_slide = 0
                for this_one in slide_data: #data[details][1] is dictionary of {slide_no:word_counts, ...}
                    if slide_data[this_one] > max_slide_wc:
                        long_slide = this_one
                        max_slide_wc = slide_data[this_one]
                string = SINGLE_PRESENTATION_ANALYSIS % (presentation, word_count, slide_count,
                                                         long_slide, max_slide_wc, round(word_count/slide_count,1))
                logging.debug("File %s (%s slides) analysis returned...", presentation, slide_count)
        else:
            logging.debug("_unpack_data() called to unpack a folder containing %s presentations.", pres_count)
            total_word_count = 0
            max_word_count = 0
            total_slide_count = 0
            verbose_pres = ''
            verbose_slide_count = 0
            for this_pres in data:
                total_word_count += data[this_pres][0]
                total_slide_count += len(data[this_pres][1])
                if data[this_pres][0] > max_word_count:
                    max_word_count = data[this_pres][0]
                    verbose_pres = this_pres
                    verbose_slide_count = len(data[this_pres][1])
            string = MULTI_PRESENTATION_ANALYSIS % (pres_count, self.file_loc.get(), total_slide_count,
                                                    total_word_count, verbose_pres, max_word_count, verbose_slide_count)
        self._update_analysis(string)
        return True

    def _update_analysis(self, wanted_string = ''):
        '''Puts analysis string on second tab. Initially sets to placeholder text.
        If passed a string, it'll put it on the analysis tab.'''
        if self.my_pres == '' and wanted_string == '':
            self.slide_analysis_str.set(ANALYSIS_PLACEHOLDER_TEXT)
        else:
            self.slide_analysis_str.set(wanted_string)
    
def _delay_type(var, dest, text):
    '''Slow type in GUI

    Usage: _delay_type(StringVariableName, DestinationWidget, TestToDisplay)'''
    var.set('')
    for each_letter in text:
        var.set(var.get() + each_letter + '')
        dest.update()
        time.sleep(0.005)


run = GUI()
run.win.mainloop()