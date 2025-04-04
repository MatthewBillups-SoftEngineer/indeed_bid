from datetime import datetime
import os
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk
from urllib.parse import urlparse

from bid_record import BidRecord
from job_bidder import JobBidder
from url_processor import URLProcessor
from utils import InitDriver
from bs4 import BeautifulSoup
from gpt_helper import ParseGPTResult, CpyGPTInstructionMsg
import win32com.client

class JobBidderUI:
    def __init__(self, root, usrname: str):
        self.usrname = usrname
        self.s_datetime = ""  # To store current date and time
        self.cur_row = 2
        
        self.exp_count_jd = 4
        
        self.chkbox_show_docx = tk.BooleanVar()  # Create a variable to store the checkbox state
        self.chkbox_show_docx.set(False)
        
        self.bid_url_processor = URLProcessor("bid_urls.xlsx")
        self.bid_url_processor.LoadFile()
        
        self.root = root
        self.root.title("Job Bidder UI")

        # Setting window size
        self.root.geometry("1100x900")
        
        # Create main frames
        self.create_editing_area()
        self.create_button_area()
        
    def create_editing_area(self):
        """
        Create the editing area with a scrollable frame
        """
        # Create the outer frame to hold the canvas and scrollbar
        container = tk.Frame(self.root, width=900, height=900)
        container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Create a canvas widget
        canvas = tk.Canvas(container, width=880)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Add a scrollbar
        scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill="y")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Create a frame inside the canvas
        self.editing_area_frame = tk.Frame(canvas, width=700)
        self.editing_area_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        # Add the frame inside the canvas
        canvas_frame = canvas.create_window((0, 0), window=self.editing_area_frame, anchor="nw")

        # Ensure resizing behavior
        def on_canvas_configure(event):
            canvas.itemconfig(canvas_frame, width=event.width)

        canvas.bind("<Configure>", on_canvas_configure)

        # **Add mouse scroll binding**
        def _on_mouse_wheel(event):
            canvas.yview_scroll(-1 * (event.delta // 120), "units")

        canvas.bind_all("<MouseWheel>", _on_mouse_wheel)  # Windows & macOS
        canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))  # Linux scroll up
        canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))  # Linux scroll down
        # Add input fields with 3-column width
        fields = [
            ("Job URL", "url_entry"),
            ("Title", "title_entry"),
            ("Company", "company_entry"),
            ("Company URL", "company_url_entry"),
            ("Job Detail", "job_detail_text"),
            ("Skills", "skills_entry"),
            ("GPT Result", "gpt_text"),
            ("Summary", "summary_text"),
            ("LCG, Inc.", "company1_text"),
            ("Inherent Technologies", "company2_text"),
            ("Silverado Technologies", "company3_text"),
            ("Mytek Network Solutions", "company4_text"),
        ]

        for i, (label_text, var_name) in enumerate(fields):
            tk.Label(self.editing_area_frame, text=label_text).grid(row=i, column=0, sticky="w", padx=5, pady=5)
            if "text" in var_name:
                setattr(self, var_name, scrolledtext.ScrolledText(self.editing_area_frame, width=100, height=5))
            else:
                setattr(self, var_name, tk.Entry(self.editing_area_frame, width=100))
            
            getattr(self, var_name).grid(row=i, column=1, columnspan=3, padx=5, pady=5, sticky="we")

        self.gpt_text.bind("<<Modified>>", self.on_gpt_change)

        # Add a horizontal line with 3 Entry widgets
        tk.Label(self.editing_area_frame, text="Additional Info:").grid(row=len(fields), column=0, sticky="w", padx=5, pady=5)

        self.entry1 = tk.Entry(self.editing_area_frame, width=30)
        self.entry1.grid(row=len(fields), column=1, padx=5, pady=5, sticky="w")

        self.entry1.delete(0, tk.END)
        self.entry1.insert(0, f"{self.exp_count_jd}")
        
        # Creating the checkbox frame
        checkbox_frame = tk.Frame(self.editing_area_frame)
        checkbox_frame.grid(row=len(fields) + 1, column=0, columnspan=4, pady=10)

        self.set_first_url()
        
    # Bind the event to detect when the content of the text widget changes
    def on_gpt_change(self, event):
        # Only call parse_html_content if there is an actual change in content
        if self.gpt_text.edit_modified():
            try:
                parse_res = ParseGPTResult(self.gpt_text.get(1.0, tk.END).strip())
                
                # Iterate through the dictionary
                for sect, res in parse_res.items():
                    if "GEN_SUMMARY" in sect:
                        self.summary_text.delete(1.0, tk.END)  # Clear any previous value
                        self.summary_text.insert(tk.END, res)  # Set the job description
                    elif "GEN_COMPANY" in sect:
                        company_names = ['LCG, INC.', 'INHERENT TECHNOLOGIES', 'SILVERADO TECHNOLOGIES', 'MYTEK NETWORK SOLUTIONS']
                        for idx, name in enumerate(company_names, start=1):
                            for company_name, company_experiences in res.items():
                                if name in company_name:
                                    attr_name = f"company{idx}_text"
                                    if getattr(self, attr_name) and isinstance(getattr(self, attr_name), scrolledtext.ScrolledText):
                                        getattr(self, attr_name).delete(1.0, tk.END)
                                        getattr(self, attr_name).insert(tk.END, "\n".join(company_experiences))
                                    break
            finally:
                self.gpt_text.edit_modified(False)  # Reset the modified flag

    def set_first_url(self):
        getattr(self, 'url_entry').delete(0, tk.END)
        if self.bid_url_processor.First():
            getattr(self, 'url_entry').insert(0, self.bid_url_processor.First()[1])
            
    def create_button_area(self):
        """
        Create the button area with different action buttons
        """
        self.button_area_frame = tk.Frame(self.root, width=300, height=900, padx=10, pady=10)
        self.button_area_frame.pack(side=tk.RIGHT, fill=tk.Y)

        self.prev_url_button = tk.Button(self.button_area_frame, text="Prev", width=12, command=self.prev_url)
        self.prev_url_button.grid(row=0, column=0, pady=10, sticky="w")
        
        self.next_url_button = tk.Button(self.button_area_frame, text="Next", width=12, command=self.next_url)
        self.next_url_button.grid(row=0, column=1, pady=10, sticky="w")
        
        self.del_button = tk.Button(self.button_area_frame, text="Del", width=25, command=self.del_url)
        self.del_button.grid(row=1, column=0, columnspan=2, pady=10)
        
        self.exist_button = tk.Button(self.button_area_frame, text="Exist", width=12, command=self.exist_url)
        self.exist_button.grid(row=2, column=1, pady=10, sticky="w")
        
        self.job_detail_button = tk.Button(self.button_area_frame, text="Job Detail", width=12, command=self.get_job_detail)
        self.job_detail_button.grid(row=2, column=0, pady=10, sticky="w")

        self.resume_button = tk.Button(self.button_area_frame, text="Copy GPT Strings to Clipboard", width=25, state=tk.ACTIVE, command=self.cpy_gptstrs_clipboard)
        self.resume_button.grid(row=3, column=0, columnspan=2, pady=10)
        
        self.resume_button = tk.Button(self.button_area_frame, text="Generate Resume", width=25, state=tk.ACTIVE, command=self.generate_resume)
        self.resume_button.grid(row=4, column=0, columnspan=2, pady=10)

        self.finalize_button = tk.Button(self.button_area_frame, text="Finalize", width=25, command=self.finalize)
        self.finalize_button.grid(row=5, column=0, columnspan=2, pady=10)
        
        self.showdocx_chkbox = tk.Checkbutton(self.button_area_frame, text="Show Docx", variable=self.chkbox_show_docx)
        self.showdocx_chkbox.grid(row=6, column=0, columnspan=2, pady=10)
        
    def prev_url(self):
        bid_url = self.bid_url_processor.Prev()
        if bid_url:
            getattr(self, 'url_entry').delete(0, tk.END)
            getattr(self, 'url_entry').insert(0, bid_url[1])
        else:
            self.set_first_url()
    def next_url(self):
        bid_url = self.bid_url_processor.Next()
        if bid_url:
            getattr(self, 'url_entry').delete(0, tk.END)
            getattr(self, 'url_entry').insert(0, bid_url[1])
        else:
            self.set_first_url()
    def del_url(self):
        if self.bid_url_processor.Cur():
            self.bid_url_processor.DelCur(self.bid_url_processor.Cur()[0])
            self.bid_url_processor.SaveFile()
        self.set_first_url()
        self.resume_button.config(state=tk.ACTIVE)
        # self.finalize_button.config(state=tk.DISABLED)
        
    def exist_url(self):
        try:
            self.chk_exist_job()
        except Exception as e:
            print(str(e))
            
    def get_job_detail(self):
        """
        Fetch job details from the provided URL
        """
        self.s_datetime = datetime.now().strftime('%Y-%m-%d %H:%M:%S')  # Include time (HH:mm:ss)
        
        # Disable all buttons except "Job Detail"
        self.resume_button.config(state=tk.ACTIVE)
        # self.finalize_button.config(state=tk.DISABLED)

        url = self.url_entry.get()
        if url:
            try:
                # Call GetJobDetail and get the job details as a dictionary
                
                self.job_bidder = JobBidder(12)
                self.job_details = self.job_bidder.GetJobDetail(url)
                
                # Populate the UI fields with the returned values
                self.title_entry.delete(0, tk.END)  # Clear any previous value
                if self.job_details["title"]:
                    self.title_entry.insert(0, self.job_details["title"])  # Set the title
                
                self.company_entry.delete(0, tk.END)  # Clear any previous value
                if self.job_details["company_name"]:
                    self.company_entry.insert(0, self.job_details["company_name"])  # Set the company name
                
                self.company_url_entry.delete(0, tk.END)  # Clear any previous value
                if self.job_details["company_url"]:
                    self.company_url_entry.insert(0, self.job_details["company_url"])  # Set the company URL

                self.job_detail_text.delete(1.0, tk.END)  # Clear any previous value
                if self.job_details["desc"]:
                    self.job_detail_text.insert(tk.END, self.job_details["desc"])  # Set the job description

                self.skills_entry.delete(0, tk.END)  # Clear any previous value
                if self.job_details["skills"]:
                    self.skills_entry.insert(0, self.job_details["skills"])  # Set the title
                    
                job_description = self.job_detail_text.get(1.0, tk.END).strip()
                CpyGPTInstructionMsg(job_description, int(self.entry1.get()), self.skills_entry.get())

            except Exception as e:
                # If an exception occurs, show a message box with the error message
                messagebox.showerror("Error", "Getting job detail failed: " + str(e))
        else:
            messagebox.showwarning("Input Error", "Please enter a valid Job URL.")
            
    def cpy_gptstrs_clipboard(self):
        job_description = self.job_detail_text.get(1.0, tk.END).strip()
        CpyGPTInstructionMsg(job_description, int(self.entry1.get()), self.skills_entry.get())
        
    def chk_exist_job(self):
        result = BidRecord().Exist(self.title_entry.get(), self.company_entry.get(), self.url_entry.get(), self.company_url_entry.get())
        # self.finalize_button.config(state=tk.DISABLED)

        if result['code'] == 0:
            self.finalize_button.config(state=tk.ACTIVE)
        elif result['code'] == 1:
            # Already added
            response = messagebox.askyesno("Confirmation", "Already exists. Do you want to create a new one?")
            if response:
                self.finalize_button.config(state=tk.ACTIVE)
            else:
                raise ValueError(f"Already exists")
        elif result['code'] == 2:
            # Same company
            response = messagebox.askyesno("Confirmation", "This is the same company. Do you want to continue?")
            if response:
                self.finalize_button.config(state=tk.ACTIVE)
            else:
                raise ValueError(f"This is the same company. I don't want continue.")
        elif result['code'] == 4:
            # Same job
            response = messagebox.askyesno("Confirmation", "This is the same job. Do you want to continue?")
            if response:
                print("Do action for same job...")
                self.finalize_button.config(state=tk.ACTIVE)
            else:
                raise ValueError(f"This is the same job. I don't want continue.")
        else:
            # Error creating resume
            messagebox.showerror("Error", "Error while creating resume.")
            raise ValueError(f"Error while creating resume.")
        
    def get_base_resume_path(self):
        # Check if the selected language in the main_lang combobox is valid
        # sel_lang = self.main_lang_combo.get()

        # if sel_lang in self.lang_lst:
        #     if not os.path.exists(f"{sel_lang.lower()}_resume.docx"):  # Check if file exists before attempting to load
        #         return "base_resume.docx"
        #     return f"{sel_lang.lower()}_resume.docx"
        # else:
        return "base_resume.docx"
        
    def generate_resume(self):
        try:
            self.chk_exist_job()
        except Exception as e:
            print(str(e))
            return
            
        result = BidRecord().Exist(self.title_entry.get(), self.company_entry.get(), self.url_entry.get(), self.company_url_entry.get())
        # self.finalize_button.config(state=tk.DISABLED)

        if result['code'] == 0:
            self.finalize_button.config(state=tk.ACTIVE)
        elif result['code'] == 1:
            pass
        elif result['code'] == 2:
            # Same company
            response = messagebox.askyesno("Confirmation", "This is the same company. Do you want to continue?")
            if response:                
                self.finalize_button.config(state=tk.ACTIVE)
            else:
                return
        elif result['code'] == 4:
            # Same job
            response = messagebox.askyesno("Confirmation", "This is the same job. Do you want to continue?")
            if response:
                print("Do action for same job...")
                self.finalize_button.config(state=tk.ACTIVE)
            else:
                return
        else:
            # Error creating resume
            messagebox.showerror("Error", "Error while creating resume.")
            return

        try:
            if not self.summary_text.get(1.0, tk.END).strip():
                raise ValueError(f"Summary text is empty")
            if not self.company1_text.get(1.0, tk.END).strip():
                raise ValueError(f"LCG, Inc. text is empty")
            if not self.company2_text.get(1.0, tk.END).strip():
                raise ValueError(f"Inherent Technologies text is empty")
            if not self.company3_text.get(1.0, tk.END).strip():
                raise ValueError(f"Silverado Technologies text is empty")
            if not self.company4_text.get(1.0, tk.END).strip():
                raise ValueError(f"Mytek Network Solutions text is empty")
            
            resume_path = self.job_bidder.GenResume(
                self.get_base_resume_path(),
                self.title_entry.get(), self.company_entry.get(), self.usrname,
                {
                    '{__Summary__}' : (None, self.summary_text.get(1.0, tk.END).strip()),
                    '{__Company1__}' : ("List Bullet", self.company1_text.get(1.0, tk.END).strip()),
                    '{__Company2__}' : ("List Bullet", self.company2_text.get(1.0, tk.END).strip()),
                    '{__Company3__}' : ("List Bullet", self.company3_text.get(1.0, tk.END).strip()),
                    '{__Company4__}' : ("List Bullet", self.company4_text.get(1.0, tk.END).strip())
                }
            )

            add_dict = {
                "Site": f"{urlparse(self.url_entry.get()).scheme}://{urlparse(self.url_entry.get()).netloc}/",
                "Title": self.title_entry.get(),
                "Company": self.company_entry.get(),
                "Job Detail": self.url_entry.get(),
                "Company Url": self.company_url_entry.get(),
                "Start": self.s_datetime,
                "End": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                "Bid Duration": "",
                "Resume": os.path.abspath(resume_path),  # Local file path for the resume
            }

            if result['code'] == 1:
                # Already added
                self.cur_row = result['no'] - 1
                self.finalize(should_remove=False)
            else:
                self.cur_row = BidRecord().AddRecord(add_dict)
            # self.resume_button.config(state=tk.DISABLED)
            self.finalize_button.config(state=tk.ACTIVE)
            
            if self.chkbox_show_docx.get():
                # Start Microsoft Word application
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = True  # Makes the Word application visible

                # Open a specific document
                doc_path = f"{os.path.abspath(resume_path)}"  # Change this to your document's path
                word.Documents.Open(doc_path)
            
        except Exception as e:
            print(f"{e}")
            # If an exception occurs, show a message box with the error message
            messagebox.showerror("Error", "Generating resume failed: " + str(e))
            # self.finalize_button.config(state=tk.DISABLED)

    def finalize(self, should_remove=True):
        try:
            BidRecord().FinalizeRecord(self.cur_row)
            self.resume_button.config(state=tk.ACTIVE)
            # self.finalize_button.config(state=tk.DISABLED)
            
            fields = [
                ("Title", "title_entry"),
                ("Company", "company_entry"),
                ("Company URL", "company_url_entry"),
                ("Job Detail", "job_detail_text"),
                ("Skills", "skills_entry"),
                ("GPT Result", "gpt_text"),
                ("Summary", "summary_text"),
                ("LCG, Inc.", "company1_text"),
                ("Inherent Technologies", "company2_text"),
                ("Silverado Technologies", "company3_text"),
                ("Mytek Network Solutions", "company4_text"),
            ]
            if should_remove:
                for i, (label_text, var_name) in enumerate(fields):
                    if getattr(self, var_name) and isinstance(getattr(self, var_name), tk.Entry):
                        getattr(self, var_name).delete(0, tk.END)
                    elif getattr(self, var_name) and isinstance(getattr(self, var_name), scrolledtext.ScrolledText):
                        getattr(self, var_name).delete(1.0, tk.END)
            
            if self.bid_url_processor.Cur():
                self.bid_url_processor.DelCur(self.bid_url_processor.Cur()[0])
                self.bid_url_processor.SaveFile()
            if should_remove:
                self.set_first_url()
        except Exception as e:
            self.finalize_button.config(state=tk.ACTIVE)

if __name__ == "__main__":
    root = tk.Tk()
    app = JobBidderUI(root, "Matthew Billups")
    root.mainloop()
