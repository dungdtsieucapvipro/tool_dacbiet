import os
import subprocess
from datetime import datetime

def get_computer_name():
    return os.getenv('COMPUTERNAME', 'UnknownPC')

def get_serial_number():
    try:
        command = 'wmic bios get serialnumber'
        result = subprocess.check_output(command, shell=True, stderr=subprocess.DEVNULL, stdin=subprocess.DEVNULL)
        serial_number = result.decode().split('\n')[1].strip()
        return serial_number
    except Exception as e:
        return "UnknownSerial"

def get_windows_product_name():
    try:
        command = 'wmic os get Caption'
        result = subprocess.check_output(command, shell=True, stderr=subprocess.DEVNULL, stdin=subprocess.DEVNULL)
        product_name = result.decode().split('\n')[1].strip()
        return product_name
    except Exception as e:
        return "UnknownProductName"

def get_windows_product_id():
    try:
        command = 'wmic os get SerialNumber'
        result = subprocess.check_output(command, shell=True, stderr=subprocess.DEVNULL, stdin=subprocess.DEVNULL)
        product_id = result.decode().split('\n')[1].strip()
        return product_id
    except Exception as e:
        return "UnknownProductID"

def get_build_version():
    try:
        command = 'wmic os get Version'
        result = subprocess.check_output(command, shell=True, stderr=subprocess.DEVNULL, stdin=subprocess.DEVNULL)
        build_version = result.decode().split('\n')[1].strip()
        return build_version
    except Exception as e:
        return "UnknownBuildVersion"

def get_installed_key():
    try:
        command = 'wmic path SoftwareLicensingService get OA3xOriginalProductKey'
        result = subprocess.check_output(command, shell=True, stderr=subprocess.DEVNULL, stdin=subprocess.DEVNULL)
        installed_key = result.decode().split('\n')[1].strip()
        return installed_key if installed_key else "NoInstalledKeyFound"
    except Exception as e:
        return "UnknownInstalledKey"

def get_oem_key():
    # OEM key is generally the same as Installed Key unless it's a Retail version, using the same command here
    return get_installed_key()

def get_oem_edition():
    try:
        command = 'wmic os get SKU'
        result = subprocess.check_output(command, shell=True, stderr=subprocess.DEVNULL, stdin=subprocess.DEVNULL)
        oem_edition = result.decode().split('\n')[1].strip()
        return oem_edition if oem_edition else "NoOEMEditionFound"
    except Exception as e:
        return "UnknownOEMEdition"

def save_to_file(computer_name, serial_number, product_name, product_id, build_version, installed_key, oem_key, oem_edition):
    # Lấy thời gian hiện tại
    now = datetime.now()
    timestamp = now.strftime("%Y%m%d_%H%M")

    # Tạo tên file với format: tên máy _ mã serial máy _ ngàytháng _ giờphút.txt
    filename = f"{computer_name}_{serial_number}_{timestamp}.txt"

    # Tạo thư mục keywin ở ổ C:
    directory = r"C:\keywin"
    if not os.path.exists(directory):
        os.makedirs(directory)

    # Tạo đường dẫn đầy đủ tới file
    filepath = os.path.join(directory, filename)

    # Ghi thông tin vào file
    with open(filepath, 'w') as file:
        file.write(f"Computer Name: {computer_name}\n")
        file.write(f"Serial Number: {serial_number}\n")
        file.write(f"Product Name: {product_name}\n")
        file.write(f"Product ID: {product_id}\n")
        file.write(f"Build Version: {build_version}\n")
        file.write(f"Installed Key: {installed_key}\n")
        file.write(f"OEM Key: {oem_key}\n")
        file.write(f"OEM Edition: {oem_edition}\n")

def main():
    computer_name = get_computer_name()
    serial_number = get_serial_number()
    product_name = get_windows_product_name()
    product_id = get_windows_product_id()
    build_version = get_build_version()
    installed_key = get_installed_key()
    oem_key = get_oem_key()
    oem_edition = get_oem_edition()

    # Lưu thông tin vào file
    save_to_file(computer_name, serial_number, product_name, product_id, build_version, installed_key, oem_key, oem_edition)

if __name__ == "__main__":
    # Chạy ngầm, không hiện cửa sổ console
    try:
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        main()
    except:
        main()
