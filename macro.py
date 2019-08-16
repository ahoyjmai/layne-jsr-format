import pyautogui # for keyboard macro control
import win32gui # to switch windows
import os
import time

pyautogui.PAUSE=0.03    # standard delay in seconds, pyautogui automatically delays this long after each 

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
                     1,1,]
        # enter subtotal menu
        pyautogui.keyDown('alt')
        pyautogui.press('a')
        pyautogui.press('b')
        pyautogui.keyUp('alt')

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

def add_formatting():
        move_down_right(1,3)
        for k in range(4):
                # move to subtotal row
                pyautogui.keyDown('ctrl')
                pyautogui.press('down')
                pyautogui.keyUp('ctrl')
                #time.sleep(0.3)
                pyautogui.press('down')
                # highlight row
                pyautogui.keyDown('shift')
                pyautogui.press('space')
                pyautogui.keyUp('shift')

                # select grey cell fill
                pyautogui.keyDown('alt')
                pyautogui.press('h')
                pyautogui.press('h')
                pyautogui.keyUp('alt')
                pyautogui.press('down')
                pyautogui.press('down')
                pyautogui.press('right')
                pyautogui.press('right')
                pyautogui.press('enter')

                # select black font color
                pyautogui.keyDown('alt')
                pyautogui.press('h')
                pyautogui.press('f')
                pyautogui.press('c')
                pyautogui.keyUp('alt')
                pyautogui.press('enter')

                pyautogui.press('down')
        # highlight row one last time
        pyautogui.keyDown('shift')
        pyautogui.press('space')
        pyautogui.keyUp('shift')

        # select grey cell fill
        pyautogui.keyDown('alt')
        pyautogui.press('h')
        pyautogui.press('h')
        pyautogui.keyUp('alt')
        pyautogui.press('down')
        pyautogui.press('down')
        pyautogui.press('right')
        pyautogui.press('right')
        pyautogui.press('enter')

        # select black font color
        pyautogui.keyDown('alt')
        pyautogui.press('h')
        pyautogui.press('f')
        pyautogui.press('c')
        pyautogui.keyUp('alt')
        pyautogui.press('enter')        

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
        for i in top_windows:
                if key.lower() in i[1].lower():
#                        print("I found",key.lower(),"in", i[1].lower())
                        #print (i)
                        win32gui.ShowWindow(i[0],5)
                        win32gui.SetForegroundWindow(i[0])
                        return True

        #for i in top_windows:
                #if key.lower() in i[1].lower() and "excel" in i[1].lower():
                        
                 #       break
                        #win32gui.ShowWindow(i[0],5)
                        #win32gui.SetForegroundWindow(i[0])
                        #return True
                        
        return False

        
                
def AUTOMATE_EXCEL_FORMATTING (completefilepath,savefilename):      # delete the default value, this is just for debugging
        print("Opening file in excel...",end="")
        filename=completefilepath.replace('/','\\')
        try:
                #print('file exists=',os.path.isfile(filename))
                #print('current dir=',os.getcwd())
                
                os.startfile(filename)
                print ("complete.")
        except:
                print ("\nError loading excel file.")
                return

        #wait 10 seconds for window to load, or you can wait for user input.
        #waiting seems to be better
        
        time.sleep(5)
        
        #confirmtext="Please wait until excel file opens. \n Then click OK to start macro. Or click cancel to stop the script.\n\nIf you need to stop the macro in case of emergency, quickly move the mouse to the far corner of your computer screen"
        #title="Start JSR Macro?"
        #proceed=pyautogui.confirm(confirmtext,title,buttons=['OK','Cancel']).lower()
        #focus_window(title)   # i'm not sure this actually focuses the pop up at all.
        print("--------------------------------------------------------------------------")
        print()
        print ("The next part of the script will activate a macro which will:")
        print ("    1) Open excel on your computer")
        print ("    2) Take over your keyboard and mouse.")
        print ("Do not touch keyboard and mouse while the macro is running.")
        print ()
        print ("If you need to stop the macro in case of emergency, quickly move the mouse to the far corner of your computer screen")
        print ()
        print ("Once excel has loaded, press 'ENTER' to run the macro")
        print ("If you want to skip the macro, press Ctrl-C")
        proceed=input()
        
        if proceed=="":
                print ()
                print ("MACRO STARTED.")
                print ("PLEASE DO NOT TOUCH MOUSE AND KEYBOARD.")
                if focus_window(savefilename):
                        time.sleep(1)
                        move_to_last_worksheet()
                        go_back_x_sheets(3)
                        for k in range(7):      #do this 7 times because there are 7 sheets that need formatting
                                print("Starting sheet #",k,"...",end="")
                                add_subtotals()
                                add_formatting()
                                move_down_right(2,0)
                                go_back_x_sheets(1)
                                print("Finished sheet.")
                                time.sleep(3)
                        time.sleep(3)
                        ctrl_s_to_save()
                        print("Finshed, you can use the keyboard and mouse now")
                        time.sleep(3)
                        alt_tab()
                else:
                        print("Could not find excel window named",savefilename,". Cancelling Macro.")
        else:
                print("You have chosen to cancel the macro.")
        print()
        print("You can use the keyboard and mouse now")
        
