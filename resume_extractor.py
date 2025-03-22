from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from utils import InitDriver
class ResumeExtractor:
    def __init__(self, driver: webdriver, gpt_url: str):
        self.driver = InitDriver()
        self.gpt_url = gpt_url
        self.scraped_data = {
            "summary": None,
            "experiences": [],
            "keywords_list": []
        }

    def scrap(self):
        try:
            # Load the GPT-generated page
            print(f"Opening URL: {self.gpt_url}")
            self.driver.get(self.gpt_url)

            # Wait until the target container is present
            print("Waiting for content to load...")
            page_source = self.driver.page_source

            # Save the HTML content into a file
            with open('page_content.html', 'w', encoding='utf-8') as file:
                file.write(page_source)
            container = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.markdown.prose.w-full.break-words.dark\\:prose-invert.light"))
            )

            # Extract the second <p> child as summary
            try:
                paragraphs = container.find_elements(By.TAG_NAME, "p")
                if len(paragraphs) >= 2:
                    self.scraped_data["summary"] = paragraphs[1].get_attribute("outerHTML")
                    print("Extracted summary.")
            except NoSuchElementException:
                print("No summary found.")

            # Extract all <ul> elements except the last one as experiences
            try:
                unordered_lists = container.find_elements(By.TAG_NAME, "ul")
                if len(unordered_lists) > 1:
                    self.scraped_data["experiences"] = [ul.get_attribute("outerHTML") for ul in unordered_lists[:-1]]
                    print("Extracted experiences.")
            except NoSuchElementException:
                print("No experiences found.")

            # Extract the last <ul>'s <li> items as keywords_list
            try:
                if unordered_lists:
                    last_list = unordered_lists[-1]
                    list_items = last_list.find_elements(By.TAG_NAME, "li")
                    self.scraped_data["keywords_list"] = [li.text for li in list_items]
                    print("Extracted keywords list.")
            except NoSuchElementException:
                print("No keywords found.")

        except Exception as e:
            print(f"Error during scraping: {e}")
        finally:
            self.driver.quit()