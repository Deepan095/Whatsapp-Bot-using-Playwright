from playwright.sync_api import (
    sync_playwright,
    TimeoutError as PlaywrightTimeoutError
)
import openpyxl

EXCEL_PATH = r"C:\Users\Logbase\Desktop\Bookx\contacts.xlsx"
SCREENSHOT_PATH = r"C:\Users\Logbase\Desktop\whatsapp_message_sent.png"

MESSAGE = "Hi, this is a Playwright automation"

with sync_playwright() as p:

    browser = p.chromium.launch(
        headless=False,
        slow_mo=300
    )

    page = browser.new_page()

    page.goto(
        "https://web.whatsapp.com/",
        wait_until="domcontentloaded",
        timeout=60000
    )

    print("WhatsApp Web opened.")
    print("Scan the WhatsApp QR Code...")

    continue_button = page.get_by_role(
        "button",
        name="Continue",
        exact=True
    )

    whatsapp_loaded = False

    for attempt in range(30):

        if continue_button.is_visible():
            continue_button.click()
            print("Popup closed automatically.")
            whatsapp_loaded = True
            break

        if page.locator("#pane-side").is_visible():
            print("WhatsApp loaded.")
            whatsapp_loaded = True
            break

        page.wait_for_timeout(1000)

    if not whatsapp_loaded:
        browser.close()
        raise RuntimeError("WhatsApp did not load within 30 seconds.")

    page.wait_for_timeout(3000)

    # Search Box
    possible_search_boxes = [
        page.get_by_role("textbox", name="Search input textbox"),
        page.get_by_role("textbox", name="Search or start a new chat"),
        page.locator('div[contenteditable="true"][data-tab="3"]'),
        page.locator('div[contenteditable="true"][role="textbox"]').first
    ]

    search_box = None

    for locator in possible_search_boxes:
        try:
            locator.wait_for(state="visible", timeout=5000)
            search_box = locator
            break
        except:
            continue

    if search_box is None:
        browser.close()
        raise RuntimeError("Search box not found.")

    print("Search box found.")

    # Open Excel
    workbook = openpyxl.load_workbook(EXCEL_PATH)
    sheet = workbook.active

    contact_name = str(sheet["A2"].value).strip()

    print("Contact:", contact_name)

    search_box.click()
    search_box.press("Control+A")
    search_box.press("Backspace")
    search_box.fill(contact_name)

    page.wait_for_timeout(3000)

    side_panel = page.locator("#pane-side")

    contact = side_panel.get_by_text(
        contact_name,
        exact=True
    ).first

    contact.click()

    print("Chat opened.")

    # Message Box
    possible_message_boxes = [
        page.get_by_role("textbox", name="Type a message"),
        page.locator(
            'footer div[contenteditable="true"][role="textbox"]'
        ),
        page.locator(
            'footer div[contenteditable="true"]'
        )
    ]

    message_box = None

    for locator in possible_message_boxes:
        try:
            locator.first.wait_for(
                state="visible",
                timeout=5000
            )
            message_box = locator.first
            break
        except:
            continue

    if message_box is None:
        browser.close()
        raise RuntimeError("Message box not found.")

    message_box.click()
    message_box.fill(MESSAGE)
    message_box.press("Enter")

    print("Message sent successfully.")

    # Wait for message to appear
    page.wait_for_timeout(2000)

    # Take screenshot
    page.screenshot(
        path=SCREENSHOT_PATH,
        full_page=False
    )

    print(f"Screenshot saved to:\n{SCREENSHOT_PATH}")

    workbook.close()

    page.wait_for_timeout(5000)

    browser.close()
    