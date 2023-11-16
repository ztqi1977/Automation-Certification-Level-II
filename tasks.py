from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
from RPA.Desktop import Desktop
import time

@task
def order_robots_from_RobotSpareBin():
    # browser.configure(slowmo=100)
    """
    Orders robtos from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the recipts and the images.
    """
    open_robot_order_website()
    orders = get_orders()
    for row in orders:
        fill_the_form(row)
    archive_receipts()

def close_annoying_modal():
    page = browser.page()
    page.click("button:text('OK')")

def open_robot_order_website():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")    

def get_orders():
    http = HTTP()
    http.download("https://robotsparebinindustries.com/orders.csv", overwrite=True)
    tables = Tables()
    orders = tables.read_table_from_csv("orders.csv")
    return orders


def wait_for_image_load(preview):
    all_imgs = preview.query_selector_all('img')
    for img in all_imgs:
        img.wait_for_element_state("visible") # only "visible" can load image complete, don't use "stable"


def store_receipt_as_pdf(order_number):
    page = browser.page()
    pdf = PDF()
    sales_results_html = page.query_selector("#receipt").inner_html()
    out_pdf = f"output/receipt_{order_number}.pdf"
    pdf.html_to_pdf(sales_results_html, out_pdf)
    return out_pdf

def screenshot_robot(order_number):
    page = browser.page()
    preview = page.query_selector("#robot-preview-image")
    wait_for_image_load(preview)
    out_png = f"output/{order_number}.png"
    preview.screenshot(path=out_png)
    return out_png

def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf = PDF()
    pdf.add_files_to_pdf(
        files=[screenshot],
        target_document=pdf_file,
        append=True
    )

    
def fill_the_form(row):
    close_annoying_modal()
    page = browser.page()
    page.select_option("#head", value=row["Head"])
    page.set_checked(f"#id-body-{row['Body']}", checked=True)
    page.get_by_label("Legs").fill(row["Legs"])
    page.fill("#address", row["Address"])
    page.click("#order")
    for i in range(100):        
        order_another = page.query_selector("#order-another")
        if order_another:
            break
        time.sleep(0.1)
        alert = page.query_selector(".alert-danger")
        if alert: # only click order if there is an alert
            page.click("#order")
        
    order_number = row["Order number"]

    screenshot = screenshot_robot(order_number)    
    pdf_file = store_receipt_as_pdf(order_number)
    embed_screenshot_to_receipt(screenshot, pdf_file)
    page.click("#order-another")
    
def archive_receipts():
    archive = Archive()
    archive.archive_folder_with_zip("output", "output/receipts.zip", include="receipt_*.pdf")
    
    
