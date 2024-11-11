import os, time, shutil, pymupdf
from colorama import Fore, Back, Style
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from datetime import datetime
from pathlib import Path

# Variables
current_year = datetime.now().year
url = "https://flosscross.com/designer"
FilesToProcess = Path("G:\\My Drive\\Cross Stitch\\fcjson files")
SaveLocation = Path("G:\\My Drive\\Cross Stitch\\Patterns not yet on Etsy\\ProcessedJsons")
DownloadLocation = Path("C:\\Users\\jolen\\Downloads")

# Launch Brower
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)
chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
driver = webdriver.Chrome(options=chrome_options)
driver.get(url)
title = driver.title
driver.implicitly_wait(2)
time.sleep(5)


def ProcessPDF(PDFFile):
    # Open the PDF document
    doc = pymupdf.open(PDFFile)

    # Iterate over each page of the document
    for page in doc:
        # Find all instances of "Produced using FlossCross" on the current page
        instances = page.search_for("Produced using FlossCross")

        # Redact each instance of "Produced using FlossCross" on the current page
        for inst in instances:
            page.add_redact_annot(inst)

        # Apply the redactions to the current page
        page.apply_redactions()

    # Save the modified document
    doc.saveIncr()

    # Close the document
    doc.close()


def ProccessJsons(FilePath, FCJsonFullpath, Filename):
    PatternFolderName = Filename.split(" Cross")[0]
    PatternFolderLocation = os.path.join(SaveLocation, PatternFolderName)
    OldFCJson = FCJsonFullpath

    # Create Folder
    print(Fore.GREEN + "Creating Folder : " + Fore.BLUE + PatternFolderLocation)
    if not os.path.exists(PatternFolderLocation):
        os.makedirs(PatternFolderLocation)

    # Move jcjson File to Appropriate folder if not already there
    print(
        Fore.GREEN
        + "Moving FCJSON : "
        + Fore.BLUE
        + Filename
        + " to "
        + PatternFolderLocation
    )
    if not Path(PatternFolderLocation + Filename).is_file():
        shutil.copy(Path(FCJsonFullpath), PatternFolderLocation)
        FCJsonFullpath = PatternFolderLocation + "\\" + Filename

    # Load File into slot 1
    print("Loading " + FCJsonFullpath + " Into FlossCross.")
    file_input = driver.find_element(
        By.XPATH,
        "/html/body/div[1]/div[1]/div/div[4]/div[1]/div/div[2]/div[3]/button[3]/span[2]/span/label/div/div/div/div/input[1]",
    )
    file_input.send_keys(FCJsonFullpath)

    time.sleep(2)

    # Click Dots
    driver.find_element(
        By.XPATH,
        "/html/body/div[1]/div[1]/div/div[4]/div[1]/div/div[2]/div[5]/button[1]",
    ).click()

    time.sleep(2)

    # Click Open PDF
    # driver.find_element(By.XPATH, "/html/body/div[5]/div/a[2]/div[2]").click()
    # /html/body/div[4]/div/a[2]/div[2]
    driver.find_element(By.LINK_TEXT, "Export to PDF").click()

    time.sleep(2)

    # Select Pattern Keeper File
    print("Selecting Pattern Keeper File")
    driver.find_element(
        By.XPATH, "/html/body/div[1]/div[1]/div/div[2]/div[2]/div[1]/div/div"
    ).click()

    time.sleep(2)

    # Select Footer
    Footer = driver.find_element(
        By.XPATH,
        "/html/body/div[1]/div[1]/div/div[2]/div[2]/div[2]/div/div[2]/div/label[2]/div/div/div/input",
    )

    time.sleep(2)

    # Edit Footer
    Footer.send_keys(Keys.CONTROL + "a")
    Footer.send_keys(Keys.DELETE)
    FooterText = f"\u00a9{current_year} Copyright littlethingsbyjoe"
    time.sleep(2)
    Footer.send_keys(FooterText)

    time.sleep(2)

    # Click Save as PDF to save PK File
    driver.find_element(
        By.XPATH, "/html/body/div[1]/div[1]/div/div[2]/div[1]/button[2]"
    ).click()

    time.sleep(2)

    # Save PK File with Correct Name
    SaveBox = driver.find_element(
        By.XPATH,
        "/html/body/div[4]/div[2]/div/div[2]/label/div/div/div/input",
    )
    SaveBox.send_keys(Keys.CONTROL + "a")
    SaveBox.send_keys(Keys.DELETE)
    SaveBox.send_keys(PatternFolderName + " - PK")
    SaveBox.send_keys(Keys.ENTER)

    time.sleep(5)

    # Move PK File to Appropriate folder
    # PDFFile = DownloadLocation + " - PK.pdf"
    shutil.move(
        Path(DownloadLocation, PatternFolderName + " - PK.pdf"), PatternFolderLocation
    )
    time.sleep(5)

    # Remove Watermark from PK File
    pdffile = os.path.join(PatternFolderLocation, PatternFolderName + " - PK.pdf")
    ProcessPDF(pdffile)

    # Untick pattern keeper
    driver.find_element(
        By.XPATH, "/html/body/div[1]/div[1]/div/div[2]/div[2]/div[1]/div/div"
    ).click()

    time.sleep(2)

    for x in range(3):
        if x == 0:
            symbolmm = "3"
            pdfappend = " - 3mm color"
        elif x == 1:
            symbolmm = "4"
            pdfappend = " - 4mm color"
        else:
            # Change Symbol Size 5mm
            symbolmm = "5"
            pdfappend = " - 5mm color"

        # Input Symbol Size
        symbolsize = driver.find_element(
            By.CSS_SELECTOR, "[aria-label='Pattern symbol size (mm)']"
        )
        symbolsize.send_keys(Keys.CONTROL + "a")
        symbolsize.send_keys(Keys.DELETE)
        symbolsize.send_keys(symbolmm)
        symbolsize.send_keys(Keys.ENTER)

        time.sleep(2)

        # Click Save as PDF
        driver.find_element(
            By.XPATH, "/html/body/div[1]/div[1]/div/div[2]/div[1]/button[2]"
        ).click()

        time.sleep(2)

        # Save Color File
        SaveBox = driver.find_element(
            By.XPATH,
            "/html/body/div[4]/div[2]/div/div[2]/label/div/div/div/input",
        )
        time.sleep(2)

        SaveBox.send_keys(Keys.CONTROL + "a")
        SaveBox.send_keys(Keys.DELETE)
        SaveBox.send_keys(PatternFolderName + pdfappend)
        SaveBox.send_keys(Keys.ENTER)

        time.sleep(5)

        shutil.move(
            Path(DownloadLocation, PatternFolderName + pdfappend + ".pdf"),
            PatternFolderLocation,
        )

        time.sleep(5)

        # Remove Watermark from Color Files
        pdffile = os.path.join(
            PatternFolderLocation, PatternFolderName + pdfappend + ".pdf"
        )
        ProcessPDF(pdffile)

    # Untick Colored Pattern
    # By.CSS_SELECTOR, "[aria-label='Colored pattern']"
    driver.find_element(By.CSS_SELECTOR, "[aria-label='Colored pattern']").click()

    for x in range(3):
        if x == 0:
            # Change Symbol size to 5mm
            symbolmm = "5"
            pdfappend = " - 5mm BW"
        elif x == 1:
            # Change Symbol Size 4mm
            symbolmm = "4"
            pdfappend = " - 4mm BW"
        else:
            # Change Symbol Size 3mm
            symbolmm = "3"
            pdfappend = " - 3mm BW"

        # Input Symbol Size
        symbolsize = driver.find_element(
            By.CSS_SELECTOR, "[aria-label='Pattern symbol size (mm)']"
        )
        symbolsize.send_keys(Keys.CONTROL + "a")
        symbolsize.send_keys(Keys.DELETE)
        symbolsize.send_keys(symbolmm)
        symbolsize.send_keys(Keys.ENTER)

        # Click Save as PDF
        driver.find_element(
            By.XPATH, "/html/body/div[1]/div[1]/div/div[2]/div[1]/button[2]"
        ).click()

        time.sleep(2)

        # Save BW File
        SaveBox = driver.find_element(
            By.XPATH,
            "/html/body/div[4]/div[2]/div/div[2]/label/div/div/div/input",
        )
        SaveBox.send_keys(Keys.CONTROL + "a")
        SaveBox.send_keys(Keys.DELETE)
        SaveBox.send_keys(PatternFolderName + pdfappend)
        SaveBox.send_keys(Keys.ENTER)

        time.sleep(5)

        shutil.move(
            Path(DownloadLocation, PatternFolderName + pdfappend + ".pdf"),
            PatternFolderLocation,
        )

        time.sleep(5)

        # Remove Watermark from 3mm File
        pdffile = os.path.join(
            PatternFolderLocation, PatternFolderName + pdfappend + ".pdf"
        )
        ProcessPDF(pdffile)

    # Click Burger Menu
    driver.find_element(
        By.XPATH, "/html/body/div[1]/div[1]/div/div[2]/div[1]/button[1]"
    ).click()

    time.sleep(2)

    # Click Back to Dashboard
    driver.find_element(By.XPATH, "/html/body/div[4]/div/a[1]").click()

    time.sleep(2)

    # Click Dots
    driver.find_element(
        By.XPATH,
        "/html/body/div[1]/div[1]/div/div[4]/div[1]/div/div[2]/div[5]/button[1]",
    ).click()

    time.sleep(2)

    # Click Export to SVG
    driver.find_element(
        By.XPATH,
        "/html/body/div[4]/div/a[3]",
    ).click()

    time.sleep(2)

    # Click Save To SVG
    driver.find_element(By.XPATH, '//*[@title="svg.saveSvg"]').click()

    SaveBox = driver.find_element(
        By.XPATH,
        "/html/body/div[4]/div[2]/div/div[2]/label/div/div/div/input",
    )

    # Save BW SVG
    SaveBox.send_keys(Keys.CONTROL + "a")
    SaveBox.send_keys(Keys.DELETE)
    SaveBox.send_keys(PatternFolderName + " - BW")
    SaveBox.send_keys(Keys.ENTER)

    time.sleep(2)

    # Move BW SVG
    shutil.move(
        Path(DownloadLocation, PatternFolderName + " - BW.svg"),
        PatternFolderLocation,
    )

    # Change SVG to Color
    driver.find_element(By.CSS_SELECTOR, "[aria-label='Color blocks']").click()

    # Click Save To SVG
    driver.find_element(By.XPATH, '//*[@title="svg.saveSvg"]').click()

    SaveBox = driver.find_element(
        By.XPATH,
        "/html/body/div[4]/div[2]/div/div[2]/label/div/div/div/input",
    )

    # Save SVG Color
    SaveBox.send_keys(Keys.CONTROL + "a")
    SaveBox.send_keys(Keys.DELETE)
    SaveBox.send_keys(PatternFolderName + " - color")
    SaveBox.send_keys(Keys.ENTER)

    time.sleep(2)

    # Move color SVG
    shutil.move(
        Path(DownloadLocation, PatternFolderName + " - color.svg"),
        PatternFolderLocation,
    )

    # Click Burger Menu
    driver.find_element(
        By.XPATH, "/html/body/div[1]/div[1]/div/div[2]/div[1]/button[1]"
    ).click()

    time.sleep(2)

    # Click Back to Dashboard
    driver.find_element(By.XPATH, "/html/body/div[4]/div/a[1]").click()

    time.sleep(2)

    # Click Dots
    driver.find_element(
        By.XPATH,
        "/html/body/div[1]/div[1]/div/div[4]/div[1]/div/div[2]/div[5]/button[1]",
    ).click()

    time.sleep(2)

    # Click Export to OXS
    driver.find_element(
        By.XPATH,
        "/html/body/div[4]/div/div[2]",
    ).click()

    time.sleep(2)

    SaveBox = driver.find_element(
        By.XPATH,
        "/html/body/div[5]/div[2]/div/div[2]/label/div/div/div/input",
    )

    time.sleep(2)

    # Save OXS File
    SaveBox.send_keys(Keys.CONTROL + "a")
    SaveBox.send_keys(Keys.DELETE)
    SaveBox.send_keys(PatternFolderName)
    SaveBox.send_keys(Keys.ENTER)

    time.sleep(2)

    # Move OXS File
    shutil.move(
        Path(DownloadLocation, PatternFolderName + ".oxs"),
        PatternFolderLocation,
    )

    # Click Erase Slot
    driver.find_element(
        By.XPATH,
        "/html/body/div[4]/div/div[4]",
    ).click()

    time.sleep(2)

    # Click confirm erase
    driver.find_element(
        By.XPATH,
        "/html/body/div[5]/div[2]/div/div[4]/button[2]",
    ).click()

    if os.path.exists(FCJsonFullpath):
        os.remove(OldFCJson)
    else:
        print("The file does not exist")

    time.sleep(5)


for filename in os.listdir(FilesToProcess):
    f = os.path.join(FilesToProcess, filename)
    # checking if it is a file
    if filename.endswith(".fcjson") and os.path.isfile(f):
        ProccessJsons(FilesToProcess, f, filename)
    


driver.close()
