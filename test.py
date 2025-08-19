import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import os
from datetime import datetime

class GoogleMapsAutomation:
    def __init__(self, starting_location="Mumbai, Maharashtra, India"):
        """
        Initialize the automation script
        
        Args:
            starting_location (str): Your residential location
        """
        self.starting_location = starting_location
        self.destination = "91 Springboard, Vikhroli, Mumbai"
        self.driver = None
        self.wait = None
        
    def setup_driver(self):
        """Setup Chrome WebDriver with appropriate options"""
        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        # You may need to specify the path to your ChromeDriver
        # service = Service('/path/to/chromedriver')  # Uncomment and specify path if needed
        
        try:
            self.driver = webdriver.Chrome(options=chrome_options)
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            self.wait = WebDriverWait(self.driver, 20)
            print("✓ Chrome WebDriver initialized successfully")
        except Exception as e:
            print(f"✗ Error setting up WebDriver: {e}")
            print("Make sure ChromeDriver is installed and in your PATH")
            raise
    
    def navigate_to_maps(self):
        """Step 1: Navigate to Google Maps"""
        try:
            print("Step 1: Navigating to Google Maps...")
            self.driver.get("https://maps.google.com")
            time.sleep(3)
            print("✓ Successfully navigated to Google Maps")
        except Exception as e:
            print(f"✗ Error navigating to Google Maps: {e}")
            raise
    
    def click_directions(self):
        """Step 2: Click on Directions button"""
        try:
            print("Step 2: Clicking on Directions...")
            
            # Try multiple possible selectors for the Directions button
            directions_selectors = [
                "button[data-value='Directions']",
                "button[aria-label*='Directions']",
                "[data-value='Directions']",
                "button:contains('Directions')",
                "#searchbox-directions",
                ".searchbox-directions"
            ]
            
            directions_clicked = False
            for selector in directions_selectors:
                try:
                    if "contains" in selector:
                        # Use XPath for text-based selection
                        element = self.wait.until(EC.element_to_be_clickable(
                            (By.XPATH, "//button[contains(text(), 'Directions')]")
                        ))
                    else:
                        element = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    
                    element.click()
                    directions_clicked = True
                    print("✓ Successfully clicked on Directions")
                    break
                except:
                    continue
            
            if not directions_clicked:
                # Alternative: Use keyboard shortcut
                print("Trying keyboard shortcut for directions...")
                self.driver.find_element(By.TAG_NAME, "body").send_keys(Keys.CONTROL + ".")
                time.sleep(2)
            
            time.sleep(3)
            
        except Exception as e:
            print(f"✗ Error clicking Directions: {e}")
            raise
    
    def enter_locations(self):
        """Steps 3-4: Enter starting location and destination"""
        try:
            print("Step 3-4: Entering locations...")
            
            # Wait for directions panel to load
            time.sleep(3)
            
            # Find input fields for starting point and destination
            # Google Maps typically has two input fields when directions are active
            input_fields = self.wait.until(EC.presence_of_all_elements_located(
                (By.CSS_SELECTOR, "input[aria-label*='Starting point'], input[aria-label*='Destination'], .tactile-searchbox-input input")
            ))
            
            if len(input_fields) < 2:
                # Try alternative selectors
                input_fields = self.driver.find_elements(By.CSS_SELECTOR, ".searchbox input, .tactile-searchbox-input input")
            
            if len(input_fields) >= 2:
                # Clear and enter starting location
                starting_input = input_fields[0]
                starting_input.clear()
                starting_input.send_keys(self.starting_location)
                time.sleep(1)
                starting_input.send_keys(Keys.ENTER)
                
                print(f"✓ Entered starting location: {self.starting_location}")
                time.sleep(2)
                
                # Clear and enter destination
                destination_input = input_fields[1]
                destination_input.clear()
                destination_input.send_keys(self.destination)
                time.sleep(1)
                destination_input.send_keys(Keys.ENTER)
                
                print(f"✓ Entered destination: {self.destination}")
                time.sleep(5)  # Wait for route calculation
                
            else:
                raise Exception("Could not find location input fields")
                
        except Exception as e:
            print(f"✗ Error entering locations: {e}")
            raise
    
    def select_first_route(self):
        """Step 5: Select the first route"""
        try:
            print("Step 5: Selecting first route...")
            
            # Wait for routes to load
            time.sleep(5)
            
            # The first route is usually selected by default in Google Maps
            # But let's try to click on it to ensure it's selected
            try:
                first_route = self.wait.until(EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "[data-trip-index='0'], .directions-travel-mode-icon, .transit-directions-route-summary")
                ))
                first_route.click()
                print("✓ Selected first route")
            except:
                print("✓ First route already selected (default)")
            
            time.sleep(3)
            
        except Exception as e:
            print(f"✗ Error selecting first route: {e}")
            raise
    
    def extract_directions(self):
        """Step 6: Extract driving instructions"""
        try:
            print("Step 6: Extracting driving instructions...")
            
            # Wait for directions to load
            time.sleep(5)
            
            directions_data = []
            
            # Try to find step-by-step directions
            step_selectors = [
                ".directions-info .section-directions-trip-secondary-text",
                ".directions-info .section-directions-trip-travel-mode-details",
                "[data-index] .section-directions-trip-secondary-text",
                ".section-directions-trip",
                ".directions-mode-step-details"
            ]
            
            steps_found = False
            for selector in step_selectors:
                try:
                    steps = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    if steps:
                        for i, step in enumerate(steps):
                            direction_text = step.text.strip()
                            if direction_text:
                                directions_data.append({
                                    'Step': i + 1,
                                    'Instruction': direction_text,
                                    'Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                })
                        steps_found = True
                        break
                except:
                    continue
            
            if not steps_found:
                # Try to get general route information
                try:
                    route_info = self.driver.find_element(By.CSS_SELECTOR, ".section-directions-trip")
                    if route_info:
                        directions_data.append({
                            'Step': 1,
                            'Instruction': f"Route from {self.starting_location} to {self.destination}",
                            'Details': route_info.text,
                            'Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        })
                except:
                    pass
            
            # If still no directions found, add basic route info
            if not directions_data:
                directions_data.append({
                    'Step': 1,
                    'Instruction': f"Navigate from {self.starting_location} to {self.destination}",
                    'Note': "Detailed step-by-step instructions could not be extracted",
                    'Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                })
            
            # Create DataFrame and save to Excel
            df = pd.DataFrame(directions_data)
            filename = f"google_maps_directions_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            df.to_excel(filename, index=False)
            
            print(f"✓ Extracted {len(directions_data)} direction steps")
            print(f"✓ Saved directions to: {filename}")
            
            return filename
            
        except Exception as e:
            print(f"✗ Error extracting directions: {e}")
            # Create a basic Excel file with error info
            error_data = [{
                'Step': 1,
                'Instruction': f"Route from {self.starting_location} to {self.destination}",
                'Error': str(e),
                'Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }]
            df = pd.DataFrame(error_data)
            filename = f"google_maps_directions_error_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            df.to_excel(filename, index=False)
            return filename
    
    def take_screenshot(self):
        """Step 7: Take a screenshot"""
        try:
            print("Step 7: Taking screenshot...")
            
            # Ensure we're viewing the full directions
            time.sleep(2)
            
            screenshot_filename = f"google_maps_screenshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            self.driver.save_screenshot(screenshot_filename)
            
            print(f"✓ Screenshot saved as: {screenshot_filename}")
            return screenshot_filename
            
        except Exception as e:
            print(f"✗ Error taking screenshot: {e}")
            raise
    
    def cleanup(self):
        """Close the browser"""
        if self.driver:
            self.driver.quit()
            print("✓ Browser closed")
    
    def run_automation(self):
        """Run the complete automation process"""
        try:
            print("Starting Google Maps Directions Automation...")
            print("=" * 50)
            
            # Setup
            self.setup_driver()
            
            # Execute steps
            self.navigate_to_maps()
            self.click_directions()
            self.enter_locations()
            self.select_first_route()
            excel_file = self.extract_directions()
            screenshot_file = self.take_screenshot()
            
            print("=" * 50)
            print("✓ Automation completed successfully!")
            print(f"📊 Directions saved to: {excel_file}")
            print(f"📸 Screenshot saved to: {screenshot_file}")
            
            # Keep browser open for a few seconds to verify
            print("\nKeeping browser open for 5 seconds for verification...")
            time.sleep(5)
            
        except Exception as e:
            print(f"✗ Automation failed: {e}")
        finally:
            self.cleanup()

def main():
    """Main function to run the automation"""
    
    # You can customize your starting location here
    starting_location = input("Enter your starting location (or press Enter for default): ").strip()
    if not starting_location:
        starting_location = "Mumbai, Maharashtra, India"  # Default location
    
    # Create and run automation
    automation = GoogleMapsAutomation(starting_location)
    automation.run_automation()

if __name__ == "__main__":
    # Installation instructions
    print("Google Maps Directions Automation Script")
    print("=" * 40)
    print("Prerequisites:")
    print("1. Install required packages:")
    print("   pip install selenium pandas openpyxl")
    print("2. Install ChromeDriver:")
    print("   - Download from: https://chromedriver.chromium.org/")
    print("   - Add to PATH or specify path in setup_driver() method")
    print("3. Make sure Google Chrome is installed")
    print("=" * 40)
    print()
    
    main()