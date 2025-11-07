from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
import pandas as pd
import urllib.parse
import os

# Load your Excel file
try:
    excel_data = pd.read_excel('Recipients data.xlsx', sheet_name='Recipients')
    print(f"✅ Loaded {len(excel_data)} contacts from Excel file")
except Exception as e:
    print(f"❌ Error loading Excel file: {e}")
    exit()

# Create a clean Chrome setup without user data directory first
options = webdriver.ChromeOptions()
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--disable-gpu')
options.add_argument('--window-size=1920,1080')
options.add_argument('--start-maximized')
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
options.add_argument('--disable-blink-features=AutomationControlled')
options.add_argument('--disable-extensions')

# Remove the user-data-dir for now to avoid conflicts
# options.add_argument('--user-data-dir=./User_Data')

try:
    print("🚀 Initializing Chrome driver...")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    print("✅ Chrome driver initialized successfully")

except Exception as e:
    print(f"❌ Error initializing Chrome driver: {e}")
    print("Trying with different approach...")

    try:
        # Try with different service setup
        service = Service(executable_path=ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        print("✅ Chrome driver initialized with alternative approach")
    except Exception as e2:
        print(f"❌ Failed to initialize Chrome driver: {e2}")
        print("Please make sure:")
        print("1. Chrome browser is installed")
        print("2. No other Chrome instances are running")
        print("3. You have internet connection")
        exit()

try:
    print("🌐 Opening WhatsApp Web...")
    driver.get('https://web.whatsapp.com')

    # Wait for user to manually log in
    print("⏳ Please scan the QR code and wait for WhatsApp Web to load completely...")
    print("⏳ Waiting for 30 seconds for you to scan QR code...")
    sleep(30)

    # Check if WhatsApp Web is loaded by looking for the sidebar
    try:
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//div[@data-testid='chat-list']"))
        )
        print("✅ WhatsApp Web loaded successfully!")
    except:
        print("⚠️ WhatsApp Web might not be fully loaded. Continuing anyway...")

    input("Press ENTER after confirming WhatsApp Web is fully loaded and chats are visible...")

    sent_count = 0
    failed_count = 0

    for i, contact in enumerate(excel_data['Contact']):
        try:
            number = str(contact).strip()
            # Get the message for each contact
            message = str(excel_data['Message'][i]).strip()

            print(f"\n📱 Processing {i + 1}/{len(excel_data)}: {number}")
            print(f"   Message: {message[:50]}..." if len(message) > 50 else f"   Message: {message}")

            # URL encode the message
            encoded_message = urllib.parse.quote(message)
            url = f"https://web.whatsapp.com/send?phone={number}&text={encoded_message}"

            print("   Opening chat...")
            driver.get(url)
            sleep(8)  # Wait for page load

            # Check for invalid number
            try:
                error_elements = driver.find_elements(By.XPATH,
                                                      "//div[contains(text(), 'invalid') or contains(text(), 'phone number') or contains(text(), 'not on WhatsApp')]")
                if error_elements:
                    print(f"   ❌ Invalid phone number or not on WhatsApp: {number}")
                    failed_count += 1
                    continue
            except:
                pass

            # Wait for message input box with multiple selectors
            message_box = None
            selectors = [
                "//div[@contenteditable='true'][@data-tab='10']",
                "//div[@contenteditable='true'][@data-tab='9']",
                "//div[@contenteditable='true'][@spellcheck='true']",
                "//div[@contenteditable='true']",
                "//div[@role='textbox']"
            ]

            for selector in selectors:
                try:
                    message_box = WebDriverWait(driver, 20).until(
                        EC.presence_of_element_located((By.XPATH, selector))
                    )
                    print("   ✅ Chat loaded successfully")
                    break
                except:
                    continue

            if not message_box:
                print(f"   ⚠️ Could not load chat for {number}")
                failed_count += 1
                continue

            # Wait for message to be pre-filled
            print("   Waiting for message to pre-fill...")
            sleep(5)

            # Method 1: Try pressing Enter key
            print("   Trying to send message...")
            try:
                actions = ActionChains(driver)
                actions.send_keys(Keys.ENTER)
                actions.perform()
                sleep(3)

                # Verify message was sent by checking if input box is empty
                if message_box.text == "":
                    print("   ✅ Message sent using Enter key")
                    sent_count += 1
                else:
                    raise Exception("Message not sent")

            except Exception as e:
                print("   ⚠️ Enter key failed, trying send button...")

                # Method 2: Try send button
                try:
                    send_button_selectors = [
                        "//button[@data-tab='11']//span[@data-icon='send']",
                        "//span[@data-icon='send']",
                        "//button[contains(@class, 'send')]",
                        "//button[@aria-label='Send']"
                    ]

                    sent = False
                    for btn_selector in send_button_selectors:
                        try:
                            send_btn = WebDriverWait(driver, 5).until(
                                EC.element_to_be_clickable((By.XPATH, btn_selector))
                            )
                            send_btn.click()
                            sleep(3)
                            print("   ✅ Message sent using send button")
                            sent_count += 1
                            sent = True
                            break
                        except:
                            continue

                    if not sent:
                        print(f"   ❌ Could not send message to {number}")
                        failed_count += 1

                except Exception as e2:
                    print(f"   ❌ Send button failed: {e2}")
                    failed_count += 1

            # Wait before next contact to avoid rate limiting
            print("   Waiting before next message...")
            sleep(5)

        except Exception as e:
            print(f"   ⚠️ Unexpected error with {number}: {str(e)}")
            failed_count += 1

        print("-" * 50)

    print(f"\n📊 Final Summary:")
    print(f"   ✅ Successfully sent: {sent_count}")
    print(f"   ❌ Failed: {failed_count}")
    print(f"   📨 Total processed: {len(excel_data)}")

except Exception as e:
    print(f"❌ Critical error: {e}")
    import traceback

    traceback.print_exc()

finally:
    # Always quit the driver
    try:
        driver.quit()
        print("✅ Browser closed successfully")
    except:
        pass