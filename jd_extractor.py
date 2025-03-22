import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException

class JDExtractor:
    def __init__(self, url: str, driver: webdriver):
        self.url = url
        self.driver = driver

    def Title(self):
        raise NotImplementedError("This method should be implemented in a subclass.")

    def Company(self):
        raise NotImplementedError("This method should be implemented in a subclass.")

    def CompanyUrl(self):
        raise NotImplementedError("This method should be implemented in a subclass.")

    def Desc(self):
        raise NotImplementedError("This method should be implemented in a subclass.")

    def Skills(self):
        raise NotImplementedError("This method should be implemented in a subclass.")
    
    def URL(self):
        return self.url
    
    def Scrap(self):
        # Scraping logic will be implemented in the subclass
        pass

class IndeedJDExtractor(JDExtractor):
    def __init__(self, url: str, driver: webdriver):
        super().__init__(url, driver)
        print(f"Opening browser and loading: {url}")
        self.driver.get(self.url)
        
        # Initialize fields based on the current page
        self.title = None
        self.company_name = None
        self.company_url = None
        self.desc_simple = None
        self.desc = None
        self.skills = None

        self.Scrap()

    def Scrap(self):
        try:
            # Wait until the job title appears
            try:
                print("Waiting for job title to load...")
                self.title = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "h2[data-testid='simpler-jobTitle']"))
                ).text
                print(f"Job Title Found: {self.title}")
            except Exception as e:
                print(f"Error scraping {self.url}: {e}")

            try:
                # Extract company name
                print("Fetching company name...")
                company_name_element = self.driver.find_element(By.XPATH, "//h2[@data-testid='simpler-jobTitle']/following-sibling::div[1]/*[1]")
                self.company_name = company_name_element.text
                print(f"Company Name: {self.company_name}")
                
                # Extract company url
                try:
                    self.company_url = company_name_element.find_element(By.TAG_NAME, "a").get_attribute("href")
                except NoSuchElementException:
                    self.company_url = None
            except Exception as e:
                print(f"Error scraping {self.url}: {e}")

            try:
                # Extract job description
                print("Fetching job description...")
                self.desc_simple = self.driver.find_element(By.ID, "jobDescriptionText").text[:200]  # Print first 200 chars
                self.desc = self.driver.find_element(By.ID, "jobDescriptionText").text
                print(f"Job Description: {self.desc_simple}...")
            except Exception as e:
                print(f"Error scraping {self.url}: {e}")
                
            try:
                # Wait for <h3> with text "Skills" to appear
                skills_elements = WebDriverWait(self.driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, "//h3[contains(text(), 'Skills')]/following-sibling::div//button[starts-with(@data-testid, '')]")))

                # Extract all relevant texts
                self.skills = [elem.text.strip() for elem in skills_elements if elem.text.strip()]

                # Check if "+ Show more" button exists and click it
                try:
                    show_more_button = self.driver.find_element(By.XPATH, "//button[contains(text(), '+ show more')]")
                    if show_more_button.is_displayed():
                        show_more_button.click()
                        time.sleep(1)  # Wait for new content to load

                        # Re-fetch elements after clicking
                        skills_elements = WebDriverWait(self.driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, "//h3[contains(text(), 'Skills')]/following-sibling::div//button[starts-with(@data-testid, '')]")))
                        self.skills = [elem.text.strip() for elem in skills_elements if elem.text.strip()]
                        
                except Exception as e:
                    # No "Show more" button found, or other error
                    print(f"Show more button not found or an error occurred: {e}")
                    pass  # Continue without any additional action if the "Show more" button isn't found
            except Exception as e:
                # No "Show more" button found, or other error
                print(f"Error finding skills list or loading skills: {e}")
                self.skills = []  # Set skills to empty list if skillsList div is missing
                
            print("Skills:", self.skills)
                    
            print("Scraping completed successfully.")
        except Exception as e:
            print(f"Error scraping {self.url}: {e}")
        finally:
            self.driver.quit()

    def Title(self):
        return self.title

    def Company(self):
        return self.company_name

    def CompanyUrl(self):
        return self.company_url

    def Desc(self):
        return self.desc

    def Skills(self):
        return self.skills
