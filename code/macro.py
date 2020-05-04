import pyautogui # for keyboard macro control
import win32gui # to switch windows
import keyboard
import os
import time
import sys

regularspeed=0.03
slowspeed=0.2
pyautogui.PAUSE=regularspeed    # standard delay in seconds, pyautogui automatically delays this long after each 

def move_to_last_worksheet():
        pyautogui.keyDown('ctrl')
        for k in range(20):
                pyautogui.press('pagedown')
        pyautogui.keyUp('ctrl')

def alt_tab():
        pyautogui.keyDown('alt')
        pyautogui.press('tab')
        pyautogui.keyUp('alt')

def go_back_x_sheets(x=1):
        pyautogui.keyDown('ctrl')
        for k in range(x):
                pyautogui.press('pageup')
        pyautogui.keyUp('ctrl')

def ctrl_s_to_save():
        pyautogui.keyDown('ctrl')
        pyautogui.press('s')
        pyautogui.keyUp('ctrl')

def move_down_right(down=0, right=3):
        pyautogui.keyDown('ctrl')
        pyautogui.press('left')
        pyautogui.press('left')
        pyautogui.press('left')
        pyautogui.press('up')
        pyautogui.press('up')
        pyautogui.press('up')
        pyautogui.keyUp('ctrl')

        for k in range(right):
                pyautogui.press('right')
        for k in range(down):
                pyautogui.press('down')
        
def add_subtotals():
        shouldcheck=[0,0,0,0,0,1,
                     1,1,1,1,1,1,
                     1,1,1,1,1,1,
                     1,1,0,1,0,0,
                     1,1,1,1,1,1,
                     1,1,1,1,0,1,
                     1,1,1,1,1,1,
                     1,1,1]
        # enter subtotal menu
        pyautogui.PAUSE=slowspeed
        pyautogui.keyDown('alt')
        pyautogui.press('a')
        pyautogui.press('b')
        pyautogui.keyUp('alt')
        pyautogui.PAUSE=regularspeed

        # set to 'sum' instead of 'count'
        pyautogui.keyDown('shift')
        pyautogui.press('tab')
        pyautogui.keyUp('shift')
        pyautogui.press('up')
        pyautogui.press('up')
        pyautogui.press('enter')

        # uncheck default total checkbox
        pyautogui.press('tab')
        pyautogui.press('space')
        
        # check all the necessary checkboxes to sum
        pyautogui.press('home')
        for x in shouldcheck:
                if x==1:
                        pyautogui.press('space')
                pyautogui.press('down')
        # finish subtotal        
        pyautogui.press('enter')

def entire_row_greyfill_blackfont():
        # selects entire row of current cell and sets fill color and font color.
        #To be used on Totals and Subtotal lines.
        #shift-space is hotkey to select row
        pyautogui.keyDown('shift')
        pyautogui.press('space')
        pyautogui.keyUp('shift')
        # gives grey fill
        pyautogui.keyDown('alt')
        pyautogui.press('h')
        pyautogui.press('h')
        pyautogui.keyUp('alt')
        pyautogui.press('down')
        pyautogui.press('down')
        pyautogui.press('right')
        pyautogui.press('right')
        pyautogui.press('enter')
        # black font color
        pyautogui.keyDown('alt')
        pyautogui.press('h')
        pyautogui.press('f')
        pyautogui.press('c')
        pyautogui.keyUp('alt')
        pyautogui.press('enter')

def add_formatting():
        pyautogui.PAUSE=slowspeed
        # open conditional formatting menu
        pyautogui.keyDown('alt')
        pyautogui.press('h')
        pyautogui.press('l')
        pyautogui.press('r')
        pyautogui.keyUp('alt')

        # edit rule (need to do this otherwise "apply" button doesn't activate and nothing changes
        pyautogui.keyDown('alt')
        pyautogui.press('e')
        pyautogui.keyUp('alt')

        # apply rules and now conditional formatting should be active.
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.PAUSE=regularspeed

def add_formatting2():
# addformatting2 is not used. it was a previous semi-functioning version of highlihgting subtotal rows, but i've replaced it with addformatting() instead.
        # These rows will receive special highlighting. Add more if more are needed.
        searchterms=["MJ Total",
                     "LJ Total",
                     "FJ Total",
                     "CJ Total",
                     "C Total",
                     "Grand Total"]

        #reset cursor position in case it's weirdly highlighted
        move_down_right(1,1)
        
        for term in searchterms:
                # ctrl-f to bring up search box and type in search term
                pyautogui.keyDown('ctrl')
                pyautogui.press('f')
                pyautogui.keyUp('ctrl')
                for i in term:
                        pyautogui.press(i)
                pyautogui.press('enter')        # enter to locate search term
                for a in range(4):                
                        pyautogui.press('escape')
                        # escape to close search box.
                        # repeat multiple times to close "no item found window"
                
                entire_row_greyfill_blackfont()
                pyautogui.press('right')         # do this to stop highlighting row
                

#set_main_window()
def windowEnumerationHandler(hwnd, top_windows):
    top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))

# MAIN FUNCTION
# focus activate excel window
def focus_window(key="jsr"):
#        if __name__ == "__main__":
        results = []
        top_windows = []
        win32gui.EnumWindows(windowEnumerationHandler, top_windows)
        key2="xls"
        key3="excel"
        for i in top_windows:
                if key.lower() in i[1].lower():
                        if key2 in i[1].lower() or key3 in i[1].lower():
                                #print("I found",key.lower(),"in", i[1].lower())
                                #print (i)
                                win32gui.ShowWindow(i[0],3)
                                try: win32gui.SetForegroundWindow(i[0])
                                except: return False
                                return True
        # if you couldn't find the full file name, drop the last 5 characters.
        # just in case different PC's drop the word "excel" from the window name
        for i in top_windows:   
                if key.lower()[:-5] in i[1].lower():
                        #print("I found",key.lower(),"in", i[1].lower())
                        #print (i)
                        win32gui.ShowWindow(i[0],3)
                        try: win32gui.SetForegroundWindow(i[0])
                        except: return False
                        return True
                        
        return False

def KEYBOARD_MACRO_START():
        time.sleep(1)
        move_to_last_worksheet()
        go_back_x_sheets(5)
        for k in range(7):      #do this 7 times because there are 7 sheets that need formatting
                #print("Starting sheet #",k,"...",end="")
                j=k+1
                sys.stdout.write("Starting sheet %s of 7... " % j)
                sys.stdout.flush()
                add_subtotals()
                add_formatting()
                #print("Finished sheet.")
                sys.stdout.write("Done \n")
                sys.stdout.flush()
                move_down_right(2,0)
                go_back_x_sheets(1)
                time.sleep(1)
        time.sleep(3)
        ctrl_s_to_save()
        print("Finished, you can use the keyboard and mouse now!")
        time.sleep(3)
        alt_tab()
        
                
def AUTOMATE_EXCEL_FORMATTING (completefilepath,savefilename):      # delete the default value, this is just for debugging
        print("Opening file in excel...",end="")
        filename=completefilepath.replace('/','\\')
        try:
                #print('file exists=',os.path.isfile(filename))
                #print('current dir=',os.getcwd())
                
                os.startfile(filename)
                print ("complete.")
                print ("Now opening new excel file.")
        except:
                print ("\nError loading excel file.")
                return

        #wait 5 seconds for window to load, or you can wait for user input.
        #waiting seems to be better
        
        time.sleep(5)
        
        print ("--------------------------------------------------------------------------")
        print ()
        print ("When you are ready, the computer will control the mouse & keyboard.")
        print ("Once you hit 'ENTER' you have 10 seconds to:")
        print ("    1) Make sure the new JSR excel file is in front.")
        print ("    2) Confirm the excel window receives keyboard commands (ex: try pressing the arrow keys to see if it responds.")
        
        print ("If you need to emergency stop the macro, rapidly move the mouse to the far corner of the computer screen")
        print ("Otherwise, avoid using keyboard and mouse for a minute while the macro is running.")
        print ()
        
        print ()
        print ("Press 'ENTER' to run the macro. Press Ctrl-C to cancel.")
        proceed=input()
        
        if proceed=="":
                focus_window(savefilename)
                print ("YOU HAVE 10 SECONDS TO MAKE SURE JSR EXCEL WINDOW IS IN FRONT AND RECEIVING KEYBOARD COMMANDS.")
                print ("Macro starts in: ",end="")
                time.sleep(1)
                for cnt in range(10):
                        sys.stdout.write(str(10-cnt)+',')
                        sys.stdout.flush()
                        time.sleep(1)

                print ()
                print ("--------------------------------------------------------------------------")
                print ("MACRO STARTED.")
                print ("PLEASE DO NOT USE MOUSE AND KEYBOARD.")
                KEYBOARD_MACRO_START()
                #if focus_window(savefilename):
                        #KEYBOARD_MACRO_START()
                #else:
                        #print("Could not find excel window named",savefilename[:-5],". Cancelling Macro.")
        else:
                print("You have chosen to cancel the macro.")
        print()
        print("You can use the keyboard and mouse now.")
        
