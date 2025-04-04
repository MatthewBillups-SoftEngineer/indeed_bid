import pyperclip
from bs4 import BeautifulSoup
from utils import format_date

def CpyGPTInstructionMsg(job_desc: str, exp_cnt_jd: int, skills: str) -> str:
    """
    Generates the full GPT instruction message by joining different sections and copies it to the clipboard.
    """
    # Class member variables
            
    keyword_fields = ["Programming Languages", "Frameworks", "Databases", "DevOps", "Cloud", "Design Patterns & Architecture", "User Experience", "Development Methodologies", "Applications", "Bonus Skills", "Other"]
    full_msg = f"""
Here's the job description.

\"{job_desc}\"

And Here's my career.
__________________________________________________ 
MYTEK NETWORK SOLUTIONS - Full Stack Developer
April 2012 - March 2015 | Scottsdale, AZ
Started as a junior developer, contributing to a range of projects, and earned a solid middle-tier income.

SILVERADO TECHNOLOGIES - Full Stack Developer
May 2015 - June 2019 | Tucson, AZ
Advanced to senior developer and took on the responsibility of mentoring a team of five junior developers while maintaining a middle-tier income.

INHERENT TECHNOLOGIES - Full Stack Developer
September 2019 - May 2022 | | Chandler, AZ
Worked as a senior developer and consultant, delivering high-quality solutions and building strong client relationships.

LCG, INC. - Senior Full Stack Developer
June 2022 - Present | | Rockville, MD
Currently working as a senior developer in a consultancy role, successfully contributing to multiple projects and driving innovation.
__________________________________________________

First, only extract all keywords concerned with development, especially {", ".join(keyword_fields)}.
Group the extracted keywords by category using commas and conjunctions and show them to me.
Let's say them as EXTRACTED_KEYWORDS

Output Format
    HTML Strings using <p>, <strong>.

Second, generate summary and experiences. use <p> and <strong> tags.
        A resume summary for a Senior Full Stack Developer that is the best match for the job description in two sentences
        Years of experience is {12}+.
        Use the extracted keywords concerned with development above and emphasize them in a resume summary in bold.
Let's say summary section as GEN_SUMMARY

    <Company_name> - use <h1> tags
        Elaborate on my experience, incorporating as many of the extracted keywords as possible.
        Include any specific architecture patterns or performance metrics.
        These sentences should reflect deep technical expertise, often quantifying improvements.        
        
        Mention the achievement of uptime with cost reduction and reducing release cycles to automate testing and deployments.
        Include {", ".join(keyword_fields)} which you can find from extracted words for these main focus.
        Include any quantifiable achievements or impacts.
        use <ul>, <li>, and <strong> tags.

Special Instruction for companies 
    Don't use "**" strings in every sentences. use <ul>, <li>, and <strong> tags.
    All your bullet points should start with an Action Verb (e.g. Managed, Implemented), describe the task, and use a metric to state the result. 
    
    Technical contents in each experience sentence should be adapted to company duration.
    MYTEK NETWORK SOLUTIONS -
        {exp_cnt_jd} sentences,
        
    SILVERADO TECHNOLOGIES -
        {exp_cnt_jd} sentences
        
    INHERENT TECHNOLOGIES -
        {exp_cnt_jd } sentences
      
    LCG, INC. -
        {exp_cnt_jd } sentences
        
Let's say above section as GEN_COMPANY

Output Format.
    <div id="resume">
        <div id="extracted_keywords">
            ...(EXTRACTED_KEYWORDS)
        </div>
        <div id="gen_summary">
            ...(GEN_SUMMARY)
        </div>
        <div id="gen_company">
            ...(GEN_COMPANY)
        </div>
    <div>
    
Dont use ** string. If you did, use <strong> instead.
Remove first <strong> tags in the head of <li> sentences.
        """
    # Output all the result in a one bracket so that it is easy to copy and surround EXTRACTED_KEYWORDS, GEN_SUMMARY, and GEN_COMPANY section by <div> tag. 
        
    # Copy the message to the clipboard
    pyperclip.copy(full_msg)
    
    # Optionally, print a message confirming that the text was copied
    print("The GPT instruction message has been copied to the clipboard.")
    
    return full_msg

def ParseGPTResult(gpt_res: str) -> dict:
    try:
        # Parse the HTML content with BeautifulSoup
        soup = BeautifulSoup(gpt_res, 'html.parser')

        # Extract text below <div id="extracted_keywords">
        extracted_keywords = soup.find(id="extracted_keywords")
        extracted_keywords_content = extracted_keywords.decode_contents() if extracted_keywords else ""

        # Extract text below <div id="gen_summary">
        gen_summary = soup.find(id="gen_summary")
        gen_summary_content = gen_summary.decode_contents() if gen_summary else ""
        
        # Extract and split <div id="gen_company"> by <h1> tags and get <li> items
        gen_company = {}
        for company_section in soup.find_all(id="gen_company"):
            company_data = {}
            for h1_tag in company_section.find_all('h1'):
                company_name = h1_tag.get_text()
                li_tags = h1_tag.find_next('ul').find_all('li')
                company_data[company_name] = [str(li) for li in li_tags]
            gen_company = company_data

        return {
            'EXTRACTED_KEYWORDS': extracted_keywords_content,
            "GEN_SUMMARY": gen_summary_content,
            'GEN_COMPANY': gen_company
        }
    except Exception as e:
        print(f"An error occurred in the ParseGPTResult method: {e}")
        return {}  # Return a general error code for any other exceptions
