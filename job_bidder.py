from datetime import datetime
import os
import re
import shutil
from urllib.parse import urlparse
from exp_manage import ExpManage
from jd_extractor import IndeedJDExtractor, JDExtractor
from resume_creator import ResumeCreator
from bid_record import BidRecord
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pyperclip
from utils import InitDriver, remove_specialchars
from resume_extractor import ResumeExtractor
from openpyxl import load_workbook
import random

class JobBidder:
    def __init__(self, y_exp: int):
        self.y_exp = y_exp
        
        # Initialize and save the driver
        self.driver = InitDriver()

    def GenResume(self, base_resume: str, title: str, company: str, name: str, replacements: dict) -> str:
        """
        Generates a resume by replacing placeholders in the base resume with the provided replacements.
        """
        # Call the ResumeCreator's Gen method using the DstResumePath() for the destination path
        dst_path = self.DstResumePath(title, company, name)
        ResumeCreator(base_resume, dst_path).Gen(replacements)
        shutil.copy2(dst_path, name + '.docx')
        
        return dst_path
    
    def GetJobDetail(self, url: str) -> dict:
        """
        Extract job details from Indeed if the URL contains 'indeed.com', otherwise return an empty dictionary.
        """

        try:
            # Check if the job already exists in the database
            BidRecord()
            extractor = JDExtractor(url, self.driver)
            if re.search(r'indeed\.com', url):
                extractor = IndeedJDExtractor(url, self.driver)  # Initialize Indeed job extractor
                
            # Collect the job details
            job_details = {
                "title": extractor.Title(),
                "company_name": extractor.Company(),
                "company_url": extractor.CompanyUrl(),
                "desc_simple": extractor.Desc()[:200],  # First 200 characters
                "desc": extractor.Desc(),
                "skills": " , ".join(extractor.Skills())
            }
            return job_details
        except Exception as e:
            print(f"An error occurred while extracting job details: {e}")
            return {}  # Return an empty dictionary in case of error
        
    def DstResumePath(self, title: str, company: str, name: str) -> str:
        """
        Creates a folder named {title}_{company}_{yyyymmdd} if it doesn't exist and returns the full path to the resume file.
        """
        # Format the folder name using title, company, and the current date
        current_date = datetime.now().strftime("%Y%m%d")
        
        folder_name = f"Resume/{remove_specialchars(title)}_{remove_specialchars(company)}_{current_date}"
        
        # Create the directory if it does not exist
        if not os.path.exists(folder_name):
            os.makedirs(folder_name)
            print(f"Folder '{folder_name}' created successfully!")

        # Return the full path to the file within the created folder
        return os.path.join(folder_name, name + ".docx")
    