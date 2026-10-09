from playwright.sync_api import sync_playwright
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.drawing.image import Image
from datetime import datetime
from time import sleep
import pyautogui
import os

# CONFIGURATION

PROFILE_PATH = r"C:\Users\RGKD\EdgePlaywrightProfile"

EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

DEV_SERVERS = {
    "MAP": "https://geomartcloud-enterprise-dev.nonprod.pge.com/map/manager/site.html#",
    "IMAGE": "https://geomartcloud-enterprise-dev.nonprod.pge.com/image/manager/site.html#",
    "GEOHUB": "https://geomartcloud-enterprise-dev.nonprod.pge.com/geohub/manager/site.html#",
    "PRINT": "https://geomartcloud-enterprise-dev.nonprod.pge.com/print/manager/site.html#",
    "HOSTED": "https://geomartcloud-enterprise-dev.nonprod.pge.com/hosted/manager/site.html#",
    "TAMI1091": "https://geomartcloud-enterprise-dev.nonprod.pge.com/tami1091/manager/site.html#",
    "MOBILE": "https://geomartcloud-enterprise-dev.nonprod.pge.com/mobile/manager/site.html#"
}

TEST_SERVERS = {
    "MAP": "https://geomartcloud-enterprise-tst.nonprod.pge.com/map/manager/site.html#",
    "IMAGE": "https://geomartcloud-enterprise-tst.nonprod.pge.com/image/manager/site.html#",
    "GEOHUB": "https://geomartcloud-enterprise-tst.nonprod.pge.com/geohub/manager/site.html#",
    "PRINT": "https://geomartcloud-enterprise-tst.nonprod.pge.com/print/manager/site.html#",
    "HOSTED": "https://geomartcloud-enterprise-tst.nonprod.pge.com/hosted/manager/site.html#",
    "TAMI1091": "https://geomartcloud-enterprise-tst.nonprod.pge.com/tami1091/manager/site.html#",
    "MOBILE": "https://geomartcloud-enterprise-tst.nonprod.pge.com/mobile/manager/site.html#"
}

os.makedirs("screenshots", exist_ok=True)

dev_results = []
test_results = []

# PLAYWRIGHT TEST EXECUTION

def test_servers(page, servers, env_name, result_list):

    for app_name, url in servers.items():

        print(f"\nTesting {env_name} - {app_name}")

        screenshot_file = (
            f"screenshots\\{env_name}_{app_name}_Machines.png"
        )

        try:

            page.goto(
                url,
                timeout=120000,
                wait_until="networkidle"
            )

            page.wait_for_timeout(5000)

            page.locator("text=Machines").first.click()

            page.wait_for_timeout(5000)

            page.bring_to_front()

            sleep(3)

            pyautogui.screenshot(screenshot_file)

            machine_details = []

            rows = page.locator("tr")

            total_rows = rows.count()

            for i in range(total_rows):

                try:

                    row = rows.nth(i)

                    cells = row.locator("td")

                    if cells.count() >= 2:

                        machine_name = (
                            cells.nth(0)
                            .inner_text()
                            .strip()
                        )

                        machine_status = (
                            cells.nth(1)
                            .inner_text()
                            .strip()
                        )

                        if ".AWS.PGE.COM" in machine_name.upper():

                            machine_details.append({
                                "name": machine_name,
                                "status": machine_status
                            })

                            print(
                                f"Machine : {machine_name}"
                            )

                            print(
                                f"Status  : {machine_status}"
                            )

                except Exception as e:

                    print(
                        f"Row Error : {e}"
                    )

            total_machine_count = len(
                machine_details
            )

            started_count = sum(
                1
                for machine in machine_details
                if machine["status"].lower() ==
                "started"
            )

            stopped_count = (
                total_machine_count -
                started_count
            )
            health_status = "RED"

            if total_machine_count == 0:

                machine_comment = "No machines found"
                health_status = "RED"

            elif total_machine_count == 1:

                if started_count == 1:

                    machine_comment = (
                        "Only 1 machine found and it is Started"
                    )

                    health_status = "YELLOW"

                else:

                    machine_comment = (
                        "Only 1 machine found and it is Stopped"
                    )

                    health_status = "RED"

            elif started_count == total_machine_count:

                machine_comment = (
                    f"All {total_machine_count} machines are Started"
                )

                health_status = "GREEN"

            elif started_count == 0:

                machine_comment = (
                    f"All {total_machine_count} machines are Stopped"
                )

                health_status = "RED"

            else:

                machine_comment = (
                    f"{started_count} of "
                    f"{total_machine_count} machines are Started"
                )

                health_status = "YELLOW"
            print("=" * 60)
            print(f"Application : {app_name}")
            print(f"Total Machines : {total_machine_count}")
            print(f"Started : {started_count}")
            print(f"Stopped : {stopped_count}")
            print(f"Health : {health_status}")
            print(f"Comment : {machine_comment}")
            print("=" * 60)

            result_list.append([
                app_name,
                "All Data Stores are validated successfully",
                url,
                machine_comment,
                health_status,
                screenshot_file
            ])

        except Exception as ex:

            result_list.append([
                app_name,
                "Validation Failed",
                url,
                "Unable to determine machine status",
                ""
            ])
with sync_playwright() as p:

    context = p.chromium.launch_persistent_context(
        user_data_dir=PROFILE_PATH,
        executable_path=EDGE_PATH,
        headless=False,
        args=["--start-maximized"]
    )
    page = context.new_page()

    test_servers(
        page,
        DEV_SERVERS,
        "DEV",
        dev_results
    )

    test_servers(
        page,
        TEST_SERVERS,
        "TEST",
        test_results
    )

    context.close()

# EXCEL REPORT CREATION

def create_excel(results, env_name):

    wb = Workbook()

    ws = wb.active
    ws.title = "Machine Health"

# Headers
    ws.append([
    "Application",
    "Validation",
    "Link",
    "Machine Status",
    "Health Status",
    "Screenshot"
    ])

# Colours
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

# Write Data
    excel_row = 2

    for row_data in results:

        ws.append(row_data)

        health_status = row_data[4]

        if health_status == "GREEN":

            fill = green_fill

        elif health_status == "YELLOW":

            fill = yellow_fill

        else:

            fill = red_fill

        for col in range(1, 6):
            ws.cell(excel_row, col).fill = fill

        screenshot_path = row_data[5]

        if screenshot_path and os.path.exists(screenshot_path):

            try:
                img = Image(screenshot_path)

                img.width = 500
                img.height = 280

                ws.add_image(
                    img,
                    f"E{excel_row}"
                )

                ws.row_dimensions[excel_row].height = 220

            except Exception as img_error:
                print(f"Image insert failed: {img_error}")

        excel_row += 1

# Auto Width
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

        ws.column_dimensions[column_letter].width = max_length + 5

    ws.column_dimensions["E"].width = 80
    ws.column_dimensions["C"].width = 80

# Save Report
    output_folder = r"C:\Users\RGKD\PGE\DevSecOps MSO - Rgkd-DevSecOps-MSO"

    report_name = os.path.join(
        output_folder,
        f"MachineHealth_{env_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )

    wb.save(report_name)

    print("\n" + "=" * 80)
    print(f"{env_name} Machine Health Test Completed")
    print("=" * 80)
    print(f"Excel Report : {report_name}")
    print("Screenshots  : Embedded into Excel")
    print("=" * 80)

    return report_name

dev_report = create_excel(
    dev_results,
    "DEV"
)

test_report = create_excel(
    test_results,
    "TEST"
)

