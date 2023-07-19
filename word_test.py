from clicknium import clicknium as cc, locator, ui
from win32com.client.dynamic import Dispatch
import subprocess
import pytest
import docx
import os

file_name_default = "Document1"
file_name = "TestTextFile.docx"
file_path = os.getcwd()

@pytest.fixture()
def open_ms_word():
# Search for WINWORD.exe and store path
    application_path = Dispatch("Word.Application").Path
    application_exe_path = application_path.replace(os.sep, os.altsep) + "/WINWORD.exe"

    process = subprocess.Popen(application_exe_path)

    yield

# Try to close word. If file were saved with new file name uses current file name   
    for name in [file_name_default, file_name]:
        variables = {"filename":name}
        if cc.is_existing(locator.winword.button_file_tab, variables):
            ui(locator.winword.button_file_tab, variables).send_hotkey("%{F4}")

# If word document wasn't saved before, close Word app without saving
    if cc.wait_appear(locator.winword.window_save_before_close_word.button_dont_save, wait_timeout=2):
        ui(locator.winword.window_save_before_close_word.button_dont_save).click()

    process.kill()


def test_create_and_save_file(open_ms_word):    
    text = "Hello, "

# Check if word opened without blank document then create it
    if cc.is_existing(locator.winword.button_create_new_document):
        ui(locator.winword.button_create_new_document).click()
        if cc.is_existing(locator.winword.new_blank_document):
            ui(locator.winword.new_blank_document).double_click()

    ui(locator.winword.edit_body).set_text(text)
    ui(locator.winword.button_file_tab, {"filename":file_name_default}).click()
    ui(locator.winword.listitem_save_as).click()
    ui(locator.winword.button_browse).click()
    ui(locator.winword.window_save_as.edit_file_name).send_hotkey("%(A){DEL}")
    ui(locator.winword.window_save_as.edit_file_name).set_text(file_path + os.sep + file_name)
    ui(locator.winword.window_save_as.button_save).click()

# Check if document with this name already exist rewrite it 
    if cc.is_existing(locator.winword.window_save_before_close_word.rewrite_file_options_window):
        ui(locator.winword.window_save_before_close_word.rewrite_file_options_window).send_hotkey("{ENTER}")

    actual_result = docx.Document(file_name).paragraphs[0].text
    assert actual_result == text