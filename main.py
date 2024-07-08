import os
import time
import pandas as pd
from concurrent.futures import ThreadPoolExecutor
from queue import Queue
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import smtplib
from email.message import EmailMessage
from apscheduler.schedulers.blocking import BlockingScheduler
import schedule

# Set the current directory and chromedriver path
current_dir = os.path.dirname(os.path.abspath(__file__))
chromedriver_path = os.path.join(current_dir, "chromedriver")

# Check if the chromedriver file exists
if not os.path.exists(chromedriver_path):
    raise FileNotFoundError(f"Chromedriver not found at path: {chromedriver_path}")

def email_alert(subject, html_content, to):
    msg = EmailMessage()
    msg.set_content("Please view this email in HTML format to see the printer status report.")
    msg.add_alternative(html_content, subtype='html')
    msg['subject'] = subject
    msg['to'] = to

    user = "vcutsprinters@gmail.com"
    msg['from'] = user
    password = ""

    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(user, password)
    server.send_message(msg)
    server.quit()

def create_driver():
    """
    Creates and configures a Chrome WebDriver instance with necessary options and settings.
    """
    chrome_options = Options()
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-infobars")
    chrome_options.add_argument("--window-size=1920x1080")

    service = Service(executable_path=chromedriver_path)
    driver = webdriver.Chrome(service=service, options=chrome_options)

    driver.minimize_window()
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

    return driver

def check_paper_status(driver):
    """
    Checks the paper status of each drawer in the printer and returns a dictionary with the status.
    """
    drawers = {"Drawer 1": "Letter", "Drawer 2": "11x17", "Drawer 3": "Letter", "Drawer 4": "N/A"}
    paper_status = {drawer: "" for drawer in drawers}

    # Mapping of image sources to statuses
    image_to_status = {
        "pap_m00.gif": "Empty",
        "pap_m04.gif": "1 Bar",
        "pap_m07.gif": "2 Bar",
        "pap_m10.gif": "3 Bar"
    }

    for drawer in drawers:
        try:
            drawer_element = driver.find_element(By.XPATH, f"//th[contains(text(), '{drawer}')]/following-sibling::td/img")
            src = drawer_element.get_attribute("src")
            img_name = src.split('/')[-1]
            status = image_to_status.get(img_name, "N/A")  # Get the last part of the src and map it to status
            paper_status[drawer] = status
        except Exception as e:
            print(f"Error checking {drawer}: {e}")
            paper_status[drawer] = "N/A"

    return paper_status

def check_toner_status(driver):
    """
    Checks the toner status of each color in the printer and returns a dictionary with the levels.
    """
    toner_levels = {color: "" for color in ["Cyan", "Magenta", "Yellow", "Black"]}

    for color in toner_levels:
        try:
            toner_element = driver.find_element(By.XPATH, f"//th[contains(text(), '{color}')]/following-sibling::td")
            toner_levels[color] = toner_element.text.strip().split('%')[0].strip() + "%"
        except:
            toner_levels[color] = "N/A"

    return toner_levels

def navigate_and_scrape(url, printer_name, address, data_queue, alert_queue):
    """
    Navigates to the printer's status page, logs in, scrapes the toner and paper status,
    and stores the results in data_queue and alert_queue.
    """
    driver = create_driver()
    driver.get(url)

    try:
        # Bypass privacy error page if it appears
        try:
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "details-button"))).click()
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "proceed-link"))).click()
        except:
            pass  # No privacy error page appeared

        # Log in to the printer's status page
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.NAME, "userID"))).send_keys("")
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.NAME, "password"))).send_keys("")
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//input[@type="submit" and @value="Log In"]'))).click()
        time.sleep(3)  # Wait for the printer status page to load completely

        # Scrape toner and paper status
        toner_levels = check_toner_status(driver)
        paper_status = check_paper_status(driver)

        # Store scraped data in the data queue
        data_queue.put({
            'Printer': printer_name,
            'Address': address,
            'Cyan': toner_levels['Cyan'],
            'Magenta': toner_levels['Magenta'],
            'Yellow': toner_levels['Yellow'],
            'Black': toner_levels['Black'],
            'Drawer 1': paper_status['Drawer 1'],
            'Drawer 2': paper_status['Drawer 2'],
            'Drawer 3': paper_status['Drawer 3'],
            'Drawer 4': paper_status['Drawer 4']
        })

        # Generate alerts for low toner or paper refill
        for drawer, status in paper_status.items():
            if status == "Empty":
                alert_queue.put(f"{printer_name} needs paper refill on {drawer}")

        for color, level in toner_levels.items():
            if level != "N/A":
                percentage = int(level.replace("%", ""))
                if percentage <= 10:
                    alert_queue.put(f"{printer_name}: {color} toner needs refill")
                elif percentage <= 20:
                    alert_queue.put(f"{printer_name}: {color} toner will need refill soon")

    finally:
        driver.quit()

# Thread-safe queues to store printer data and alerts
data_queue = Queue()
alert_queue = Queue()

# Addresses of printers
addresses = [
    "509 N. 12th St", "509 N. 12th St", "325 N. Harrison St", "325 N. Harrison St", "410 N. 12th St", "410 N. 12th St",
    "814 W Franklin St", "900 E Leigh St", "901 W. Main St", "301 W Main St", "1100 E. Leigh St", "301 W Main St",
    "1201 E Marshall St", "1110 E Broad St"
]

# URLs and names of printers
printer_info = [
    ("https://hsl.1stfloor.ptr.vcu.edu:8443/", "HSLFirstFloor"), ("https://hsl.2ndfloor.ptr.vcu.edu:8443/", "HSLSecondFloor"),
    ("https://oneprint.pollak4.ptr.vcu.edu:8443/rps/", "PollakFloor4"), ("https://oneprint.pollak3.ptr.vcu.edu:8443/rps/", "PollakFloor3"),
    ("https://smith.130.ptr.vcu.edu:8443/", "SmithRoom130"), ("https://smith.350.ptr.vcu.edu:8443/", "SmithRoom350"),
    ("https://franklint.101.ptr.vcu.edu:8443/rps/", "Frank101"), ("https://chp.6030.ptr.vcu.edu:8443/", "CHPRoom6030"),
    ("https://temple.1143.ptr.vcu.edu:8443/", "TempleRoom1143"), ("https://harris.5116.ptr.vcu.edu:8443/rps/", "HarrisFloor5Room5116"),
    ("https://2gu00623.sonb2.ptr.vcu.edu:8443/", "SONB2Floor2Room623"), ("https://2jg04803.sndh2.ptr.vcu.edu:8443/", "SNDH2Floor2Room4803"),
    ("https://4bt04340.mmec6.ptr.vcu.edu:8443/", "MMECFloor6"), ("https://2gu01354.hsct2.ptr.vcu.edu:8443/rps/", "HuntonFloor2")
]

def check_printers():
    # Use ThreadPoolExecutor to run multiple instances of the function concurrently
    with ThreadPoolExecutor(max_workers=14) as executor:
        futures = [executor.submit(navigate_and_scrape, url, name, address, data_queue, alert_queue) for (url, name), address in zip(printer_info, addresses)]
        for future in futures:
            future.result()

    # Collect data from queues
    printer_data = []
    alerts = []

    while not data_queue.empty():
        printer_data.append(data_queue.get())

    while not alert_queue.empty():
        alerts.append(alert_queue.get())

    # Create a DataFrame from the printer data
    df = pd.DataFrame(printer_data)

    # Convert toner levels to numeric for sorting purposes, handle "N/A" cases
    for color in ['Cyan', 'Magenta', 'Yellow', 'Black']:
        df[color] = pd.to_numeric(df[color].str.replace('%', ''), errors='coerce').fillna(100).astype(int)

    # Sort DataFrame by toner and paper levels
    df.sort_values(by=['Cyan', 'Magenta', 'Yellow', 'Black', 'Drawer 1', 'Drawer 2', 'Drawer 3', 'Drawer 4'], inplace=True)

    # Convert toner levels back to string with percentage sign
    for color in ['Cyan', 'Magenta', 'Yellow', 'Black']:
        df[color] = df[color].astype(str) + '%'

    # Create the HTML report with VCU-themed styling
    html_content = """
    <html>
    <head>
        <title>Printer Status</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 0; padding: 0; background-color: #f8f9fa; }
            .header { background-color: #FEC52E; color: #000000; padding: 20px; text-align: center; }
            .container { padding: 20px; }
            table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
            th, td { border: 1px solid #dddddd; text-align: left; padding: 8px; }
            th { background-color: #FEC52E; color: #000000; }
            .cyan { background-color: cyan; color: black; }
            .magenta { background-color: magenta; color: black; }
            .yellow { background-color: #FEC52E; color: #000000; }
            .black { background-color: black; color: white; }
            tr:nth-child(even) { background-color: #f2f2f2; }
            .alert { color: red; font-weight: bold; }
            .footer { background-color: #000000; color: #FFFFFF; padding: 10px; text-align: center; position: fixed; bottom: 0; width: 100%; }
            .low-toner { color: red; }
            .empty-paper { color: red; }
        </style>
    </head>
    <body>
        <div class="header"><h1>VCU Technology Services Printer Status Report</h1></div>
        <div class="container">
        <table class='printer-table'><thead><tr><th>Printer</th><th>Address</th><th class='cyan'>Cyan</th><th class='magenta'>Magenta</th><th class='yellow'>Yellow</th><th class='black'>Black</th><th>Drawer 1</th><th>Drawer 2</th><th>Drawer 3</th><th>Drawer 4</th></tr></thead><tbody>
    """

    for _, row in df.iterrows():
        html_content += "<tr>"
        html_content += f"<td>{row['Printer']}</td>"
        html_content += f"<td>{row['Address']}</td>"
        for color in ['Cyan', 'Magenta', 'Yellow', 'Black']:
            level = int(row[color].replace('%', ''))
            html_content += f"<td class='low-toner'>{row[color]}</td>" if level <= 10 else f"<td>{row[color]}</td>"
        for drawer in ['Drawer 1', 'Drawer 2', 'Drawer 3', 'Drawer 4']:
            status = row[drawer]
            html_content += f"<td class='empty-paper'>{status}</td>" if status == "Empty" else f"<td>{status}</td>"
        html_content += "</tr>"

    html_content += "</tbody></table>"

    # Add alerts if there are any
    if alerts:
        html_content += "<h2>Alerts</h2><ul class='alert'>"
        for alert in alerts:
            html_content += f"<li>{alert}</li>"
        html_content += "</ul>"
    else:
        html_content += "<p>No alerts. All printers are in good condition.</p>"

    html_content += "</div></body></html>"

    # Send email with the HTML report content
    subject = "VCU Technology Services Printer Status Report"
    email_alert(subject, html_content, "luangvithamsp@vcu.edu")

    return alerts

# To keep track of new alerts
previous_alerts = set()

def run_job():
    global previous_alerts
    new_alerts = check_printers()
    new_alerts_set = set(new_alerts)
    
    # Check if there are new alerts
    if new_alerts_set - previous_alerts:
        for alert in new_alerts_set - previous_alerts:
            email_alert("New Printer Alert", f"<p>{alert}</p>", "luangvithamsp@vcu.edu")
    
    # Update previous alerts
    previous_alerts = new_alerts_set

# Schedule the job to run every 2 hours
scheduler = BlockingScheduler()
scheduler.add_job(run_job, 'interval', hours=1)

try:
    scheduler.start()
except (KeyboardInterrupt, SystemExit):
    pass
