import win32gui,  win32com.client,win32con, pyautogui, win32api
from time import sleep, time
from playsound import playsound
pyautogui.PAUSE = 0
pyautogui.FAILSAFE = False

def callback(hwnd, helper):

    winname = win32gui.GetWindowText(hwnd)
    if helper.win_name in winname and not "Visual" in winname:
        if not helper.use_balabolka or len(winname) > len(helper.win_name):
            print("Balabolka Window found: %s " % win32gui.GetWindowText(hwnd))

            helper.driver_win_handle = hwnd
    elif "vrchat" == winname.lower():
        print("VRChat Window found: %s " % win32gui.GetWindowText(hwnd))

        helper.vr_chat_handle = hwnd
    elif "VRCHAT" in  winname:
        print("Window %s: " % win32gui.GetWindowText(hwnd))

        helper.this_win_handle = hwnd


res = pyautogui.size()
small_win = [2560, 200, 600, 250]
big_win = [2560, 400, 1800, 700]
def_win = [res.width//2, 100, res.width//2, res.height//2]

class helper:
    def __init__(self) -> None:
        self.window_pos_x, self.window_pos_y, self.window_size_x,self.window_size_y = def_win

        self.use_immersive_reader = True
        self.background_mode_activated = False
        self.prev_log = ""
        self.use_balabolka = True
        self.driver_win_handle = 0
        self.vr_chat_handle = 0
        self.this_win_handle = 0
        self.status_active = 0
        self.past_phrases = ["" for x in range(10)]
        self.curr_phrase_selector = 9
        self.font_size = 22
        self.win_name = "immersive" if not self.use_balabolka else "Balabolka"
        self.input_desired_X = 40

        self.use_keylogger_for_input = False
        self.speak_task_queued = False

        win32gui.EnumWindows(callback, self)
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('%')

    def speak_with_word(self):
        print("\n Speaking with word")
        pyautogui.press('home')  
        pyautogui.keyDown('up') 
        pyautogui.keyDown('ctrl')  
        pyautogui.press('space')    
        pyautogui.keyUp('ctrl')  

    def speak_with_balabolka(self):
        print("\n Speaking with balabolka")

        pyautogui.press('f5')  

    def focus_speaker(self):
        print("\nFocusing balabolka window")
        rect = win32gui.GetWindowRect(self.driver_win_handle)
        pyautogui.click(rect[0] + 30, rect[1] + 30)
        win32gui.ShowWindow(self.driver_win_handle, win32con.SW_MAXIMIZE)
        win32gui.MoveWindow(self.driver_win_handle, self.window_pos_x, self.window_pos_y, self.window_size_x, self.window_size_y, True)
        win32gui.SetForegroundWindow(self.driver_win_handle)  #try

    def focus_vrchat(self):
        print("\nFocusing VRchat window")
        rect = win32gui.GetWindowRect(self.vr_chat_handle)
        win32gui.ShowWindow(self.vr_chat_handle, win32con.SW_MAXIMIZE)
        pyautogui.click(rect[0] + 30, rect[1] + 30)
        win32gui.SetForegroundWindow(self.vr_chat_handle)

if __name__ == "__main__":
    h = helper()
    try:
        win32gui.ShowWindow(h.driver_win_handle, win32con.SW_MAXIMIZE)

        win32gui.MoveWindow(h.driver_win_handle, h.window_pos_x, h.window_pos_y, h.window_size_x, h.window_size_y, True)
    except Exception as e:
        print("error ", e, " is balabolka running?")
        
    while 1:
        sleep(0.01)

        if win32api.GetAsyncKeyState(114) == -32767:
            h.status_active = 1 - h.status_active
            if h.status_active:
                try:
                    h.focus_speaker()
                except Exception as e:
                    win32gui.EnumWindows(callback, h)
                    try:
                        h.focus_speaker()
                    except:
                        print("failed to get balabolka window, is it running?")
                    print(e)
                pyautogui.keyDown('ctrl')  
                pyautogui.press('a')    
                pyautogui.keyUp('ctrl')
                pyautogui.press('backspace')
                try:    
                    playsound('beepwav.wav')
                except:
                    print("failed to play beep sound")


            else:
                if h.use_balabolka:
                    h.speak_with_balabolka()
                else:
                    h.speak_with_word()
                try:
                    h.focus_vrchat()
                except Exception as e:
                    win32gui.EnumWindows(callback, h)
                    try:
                        h.focus_vrchat()
                    except:
                        print("failed to get vr chat window, is it running?")
                        pyautogui.click(50, 50)

                    print(e)




    