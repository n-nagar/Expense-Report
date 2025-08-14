from selenium import webdriver
import json
import base64
import os

def html_to_pdf_chrome(html_path, pdf_path):
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--headless=new")  # headless mode
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--kiosk-printing")

    driver = webdriver.Chrome(options=chrome_options)

    file_url = "file://" + os.path.abspath(html_path)
    driver.get(file_url)

    # Tell Chrome to print to PDF via DevTools
    result = driver.execute_cdp_cmd(
        "Page.printToPDF",
        {
            "printBackground": True,  # keep CSS background colors
            "landscape": False,
            "scale": 1,
            "paperWidth": 8.27,  # A4
            "paperHeight": 11.69,  # A4
        }
    )

    pdf_data = base64.b64decode(result['data'])
    with open(pdf_path, "wb") as f:
        f.write(pdf_data)

    driver.quit()
    print(f"✅ PDF saved to: {pdf_path}")

if __name__ == "__main__":
    html_to_pdf_chrome("uber.html", "uber.pdf")
