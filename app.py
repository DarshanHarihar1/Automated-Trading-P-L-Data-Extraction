from AppOpener import open
import pyautogui
import time
import pytesseract
from pywinauto.application import Application
import pandas as pd
from datetime import datetime


pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

file_path = ""


def initialize_output_file():
    """
    Creates a csv file with name as the current date
    """
    global file_path 
    current_date = datetime.now().strftime('%Y-%m-%d')
    file_name = f"{current_date}.csv"
    file_path = f"{file_name}"
    headers = ['login', 'output']
    empty_df = pd.DataFrame(columns=headers)
    empty_df.to_csv(file_path, index=False)

def login_to_account(login, password, server):
    """
    Uses user creds to log in to the MT4 app
    """
    pyautogui.hotkey('shift', 'tab')
    pyautogui.write(login)
    pyautogui.press('tab')
    pyautogui.write(password)
    pyautogui.press('tab')
    pyautogui.write(server)
    pyautogui.press('enter')
    time.sleep(1)

def select_button(image_path):
    """
    Function to select the button given by an image
    """
    while True:
        try:
            location = pyautogui.locateOnScreen(image_path, confidence= 0.7)
            if location is not None:
                pyautogui.click(location)
                return location
            else:
                time.sleep(0.5)
        except Exception as e:
            print("Error", e)

def get_pl_data_img(x, y):
    """
    Function to extract the P&L Data as an image
    """
    pyautogui.moveTo(x, y)
    pyautogui.hotkey('ctrl', 'end')
    try:
        pl_location = pyautogui.locateOnScreen("pl.PNG", confidence= 0.7)
        
        if pl_location is not None:
            x,y = pyautogui.locateCenterOnScreen("pl.PNG", confidence= 0.7)
            im = pyautogui.screenshot(region=(int(pl_location.left), int(pl_location.top), int(pl_location.width)*5, int(pl_location.height)))
            return im
    except Exception as e:
        print("Error", e)

def extract_pl_from_image(im):
    """
    Function to extract string from the P&L Image
    """
    text = pytesseract.image_to_string(im, config="--psm 6")
    output = float(text.split("Credit")[0].split(":")[1].strip())
    return output

def save_to_file(login, output):
    """
    Saves P&L data to the output csv file
    """
    global file_path 
    new_rows = [
        {'login': login, 'output': output}
    ]

    existing_df = pd.read_csv(file_path)
    new_rows_df = pd.DataFrame(new_rows)
    updated_df = pd.concat([existing_df, new_rows_df], ignore_index=True)
    updated_df.to_csv(file_path, index=False)

def main():

    initialize_output_file()
    user_data = pd.read_csv("user_data.csv")

    for _, row in user_data.iterrows():

        open("Exness Technologies MT4", match_closest=True)
        time.sleep(2.5)

        login = str(row['login']) 
        password = str(row['password']) 
        server = str(row['server'])
        login_to_account(login, password, server)
        time.sleep(1)

        acchistory_loc = select_button("account_history.PNG")
        time.sleep(0.5)

        random_region_x = acchistory_loc.left
        random_region_y = acchistory_loc.top - acchistory_loc.height*4
        pyautogui.click(random_region_x, random_region_y, button='right')

        time.sleep(0.5)

        custom_period_loc = select_button("custom_period.PNG")

        time.sleep(0.5)

        app = Application().connect(title = "Custom period")
        period_combo_box = app.window(title="Custom period").ComboBox
        period_combo_box.select("Last week")
        pyautogui.press('enter')

        time.sleep(2)

        im = get_pl_data_img(random_region_x, random_region_y)
        pl = extract_pl_from_image(im)
        save_to_file(login, pl)

        pyautogui.hotkey('alt','f4')

        time.sleep(0.5)

if __name__ == "__main__":
    main()