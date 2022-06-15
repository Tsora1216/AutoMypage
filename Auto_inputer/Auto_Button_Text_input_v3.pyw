import tkinter
import pyautogui
import time
import pyperclip
import openpyxl
import os

#タブを有効化無効化
tab_on_flag=1
def tab_on_off(event):
    if tab_on_flag==1:
        tab_on_flag=0
        event.widget.config(fg="break")
    elif tab_on_flag==0:
        tab_on_flag=1
        event.widget.config(fg="red")


def button_click(event):
    event.widget.config(fg="red")
    input_text=event.widget.cget("text")
    pyperclip.copy(input_text)
    pyautogui.hotkey("altleft","tab")
    time.sleep(0.1)
    #pyperclip.paste()
    pyautogui.hotkey("ctrl","v")
    if tab_on_flag==1:
        pyautogui.hotkey("tab")
    time.sleep(0.1)
    pyautogui.hotkey("altleft","tab")
    #print(input_text)
    
# ウィンドウ作成
root = tkinter.Tk()
# ウィンドウサイズ
root.geometry("380x1000")
root.title('Auto_Text_Inputer')
root.attributes("-topmost", True)

# ボタン作成と配置
if tab_on_flag==1:
    tab_text="自動タブ無効化"
elif tab_on_flag==0:
    tab_text="自動タブ有効化"
    
tab_button = tkinter.Button(root, text=tab_text, height=2, width=10)
tab_button.bind("<ButtonPress>", button_click)
tab_button.grid(row=1, column=1, padx=5, pady=5)

wb = openpyxl.load_workbook("ES_input.xlsx")
wss = wb.sheetnames

for i in range(len(wss)):
    ws = wb[wss[i]]
    for row in ws.rows:
        for c in row:
            try:
                if(c.value[0]=="."):
                    label = tkinter.Label(root, text = c.value[1:], bg ='LightSteelBlue1')
                    label.grid(row=c.row+1, column=c.column, padx=2, pady=2)
                    continue
            except TypeError:
                None
            name = tkinter.Button(root, text=c.value, height=2, width=10)
            name.bind("<ButtonPress>", button_click)
            name.grid(row=c.row+1 ,column=c.column, padx=2, pady=2)#タブボタン用に一行下げ


# メインループ
root.mainloop()

