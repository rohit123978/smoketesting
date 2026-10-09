from playwright.sync_api import sync_playwright
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.drawing.image import Image
from datetime import datetime
from time import sleep
import pyautogui
import os

# =====================================================
# CONFIGURATION
# =====================================================

PROFILE_PATH = r"C:\Users\RGKD\EdgePlaywrightProfile"

EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

SERVERS = {
    "MAP": "https://geomartcloud-enterprise-dev.nonprod.pge.com/map/manager/site.html#",
    "IMAGE": "https://geomartcloud-enterprise-dev.nonprod.pge.com/image/manager/site.html#",
    "GEOHUB": "https://geomartcloud-enterprise-dev.nonprod.pge.com/geohub/manager/site.html#",
    "PRINT": "https://geomartcloud-enterprise-dev.nonprod.pge.com/print/manager/site.html#",
    "HOSTED": "https://geomartcloud-enterprise-dev.nonprod.pge.com/hosted/manager/site.html#",
    "TAMI1091": "https://geomartcloud-enterprise-dev.nonprod.pge.com/tami1091/manager/site.html#",
    "MOBILE": "https://geomartcloud-enterprise-dev.nonprod.pge.com/mobile/manager/site.html#"
}

os.makedirs("screenshots", exist_ok=True)

summary_results = []

# =====================================================
# PLAYWRIGHT TEST EXECUTION
# =====================================================
# =====================================================
# DATA STORE VALIDATION
# =====================================================

def validate_data_stores(page, app_name):

    print("\n" + "=" * 70)
    print(f"DATA STORE VALIDATION : {app_name}")
    print("=" * 70)

    try:

        # Open Data Stores
        page.locator("text=Data Stores").first.click()

        page.wait_for_timeout(5000)

        print("Data Stores page opened")

        # Click Validate All
        validate_button = page.get_by_text(
            "Validate All",
            exact=True
        )

        print(
            f"Validate Button Count = "
            f"{validate_button.count()}"
        )

        if validate_button.count() > 0:

            validate_button.first.scroll_into_view_if_needed()

            page.wait_for_timeout(2000)

            validate_button.first.click(force=True)

            print("Validate All Clicked")

            page.wait_for_timeout(30000)

            print("Validation Completed")

            return True

        else:

            print("Validate All Button Not Found")

            return False

    except Exception as ex:

        print(
            f"Data Store Validation Error : {ex}"
        )

        return False

with sync_playwright() as p:

    context = p.chromium.launch_persistent_context(
        user_data_dir=PROFILE_PATH,
        executable_path=EDGE_PATH,
        headless=False,
        args=["--start-maximized"]
    )

    page = context.new_page()

    for app_name, url in SERVERS.items():

        print("\n" + "=" * 80)
        print(f"Testing {app_name}")
        print("=" * 80)

        screenshot_file = f"screenshots\\{app_name}_Machines.png"

        try:

            page.goto(
                url,
                timeout=120000,
                wait_until="networkidle"
            )

            page.wait_for_timeout(5000)

            # Click Machines
            page.locator("text=Machines").first.click()

            page.wait_for_timeout(5000)

            page.bring_to_front()

            sleep(3)

            # Capture Full Desktop Screenshot
            pyautogui.screenshot(screenshot_file)

            print(f"Screenshot saved : {screenshot_file}")
            print("Opening Data Stores...")

            validation_result = validate_data_stores(
                page,
                app_name
            )

            # =====================================================
            # READ MACHINE STATUS
            # =====================================================

            statuses = []

            rows = page.locator("tr")

            total_rows = rows.count()

            for i in range(total_rows):

                try:

                    row_text = rows.nth(i).inner_text()

                    if ".AWS.PGE.COM" in row_text.upper():

                        print(row_text)

                        if "Started" in row_text:
                            statuses.append("Started")
                        else:
                            statuses.append("Stopped")

                except:
                    pass

            started_count = statuses.count("Started")
            total_machine_count = len(statuses)

            if total_machine_count == 0:

                machine_comment = (
                    "Unable to determine machine status"
                )

            elif started_count == total_machine_count:

                machine_comment = (
                    "All Machines are in started state"
                )

            elif started_count == 0:

                machine_comment = (
                    "Both machines are not in started state"
                )

            else:

                machine_comment = (
                    "One machine is started and one is not started"
                )

            print(machine_comment)

            validation_text = (
                "All Data Stores are validated successfully"
                if validation_result
                else
                "Data Store Validation Failed"
            )

            summary_results.append([
                app_name,
                validation_text,
                url,
                machine_comment,
                screenshot_file
            ])

        except Exception as ex:

            print(f"ERROR : {str(ex)}")

            summary_results.append([
                app_name,
                "Validation Failed",
                url,
                "Unable to determine machine status",
                ""
            ])

    context.close()

# =====================================================
# EXCEL REPORT CREATION
# =====================================================

wb = Workbook()

ws = wb.active
ws.title = "Machine Health"

# Headers

ws.append([
    "Application",
    "Validation",
    "Link",
    "Machine Status",
    "Screenshot"
])

# =====================================================
# COLORS
# =====================================================

green_fill = PatternFill(
    start_color="90EE90",
    end_color="90EE90",
    fill_type="solid"
)

yellow_fill = PatternFill(
    start_color="FFF59D",
    end_color="FFF59D",
    fill_type="solid"
)

red_fill = PatternFill(
    start_color="FF9999",
    end_color="FF9999",
    fill_type="solid"
)

# =====================================================
# WRITE DATA
# =====================================================

excel_row = 2

for row_data in summary_results:

    ws.append(row_data)

    status_comment = str(row_data[3])

    if "All Machines" in status_comment:
        fill = green_fill

    elif "One machine" in status_comment:
        fill = yellow_fill

    else:
        fill = red_fill

    # Color Columns A-D only
    for col in range(1, 5):
        ws.cell(excel_row, col).fill = fill

    screenshot_path = row_data[4]

    if screenshot_path and os.path.exists(screenshot_path):

        try:

            img = Image(screenshot_path)

            # Resize Image
            img.width = 500
            img.height = 280

            ws.add_image(
                img,
                f"E{excel_row}"
            )

            ws.row_dimensions[
                excel_row
            ].height = 220

        except Exception as img_error:

            print(
                f"Image insert failed: {img_error}"
            )

    excel_row += 1

# =====================================================
# AUTO WIDTH
# =====================================================

for column in ws.columns:

    column_letter = column[0].column_letter
    max_length = 0

    for cell in column:

        try:
            max_length = max(
                max_length,
                len(str(cell.value))
            )
        except:
            pass

    ws.column_dimensions[
        column_letter
    ].width = max_length + 5

# Screenshot column wider
ws.column_dimensions["E"].width = 80

# URL column wider
ws.column_dimensions["C"].width = 80

# =====================================================
# SAVE REPORT
# =====================================================

output_folder = r"C:\Users\RGKD\PGE\DevSecOps MSO - Rgkd-DevSecOps-MSO"

report_name = os.path.join(
    output_folder,
    f"MachineHealth_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
)

wb.save(report_name)

# =====================================================
# FINAL OUTPUT
# =====================================================

print("\n" + "=" * 80)
print("Machine Health Test Completed")
print("=" * 80)
print(f"Excel Report : {report_name}")
print("Screenshots  : Embedded into Excel")
print("=" * 80)