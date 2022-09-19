import autopyBot, win32gui,  win32com.client,win32con, pyautogui, win32api
path = r"imgs"
p = autopyBot.autopy.autopy(path)
import pygetwindow as gw
from time import sleep, time
from playsound import playsound
pyautogui.PAUSE = 0
pyautogui.FAILSAFE = False
# import pyttsx3
# engine = pyttsx3.init()
# #engine.say("I will speak this text")
# voices = engine.getProperty("voices")
# print(voices)
# engine.setProperty("voice", voices[3].id)
# engine.save_to_file("faster text to speech", 'test.mp3')

# engine.runAndWait()
# sleep(1)
# playsound("test.mp3") 

#ret = p.find([p.i.play], loop=2, timeout=10, region=[2560, 200, 500, 200])
#p.mouse_move(ret.found, 0, 0)

#from gtts import gTTS 
# myobj = gTTS(text="mytext", lang='en-uk', slow=False)
# myobj.save("welcome.mp3")
# playsound("welcome.mp3") 

def callback(hwnd, helper):

    rect = win32gui.GetWindowRect(hwnd)
    x = rect[0]
    y = rect[1]
    w = rect[2] - x
    h = rect[3] - y
    winname = win32gui.GetWindowText(hwnd)
    if helper.win_name in winname and not "Visual" in winname:
        if not helper.use_balabolka or len(winname) > len(helper.win_name):
            print("Window %s:" % win32gui.GetWindowText(hwnd))
            print("\tLocation: (%d, %d)" % (x, y))
            print("\t    Size: (%d, %d)" % (w, h))
            helper.driver_win_handle = hwnd
    elif "vrchat" == winname.lower():
        print("Window %s:" % win32gui.GetWindowText(hwnd))
        print("\tLocation: (%d, %d)" % (x, y))
        print("\t    Size: (%d, %d)" % (w, h))
        helper.vr_chat_handle = hwnd
    elif "VRCHAT" in  winname:# == "VRCHAT tts help":
        print("Window %s:" % win32gui.GetWindowText(hwnd))
        print("\tLocation: (%d, %d)" % (x, y))
        print("\t    Size: (%d, %d)" % (w, h))
        helper.this_win_handle = hwnd
   # elif (len(winname)):
   #\     print("Window %s:" % win32gui.GetWindowText(hwnd))

res = pyautogui.size()
small_win = [2560, 200, 600, 250]
big_win = [2560, 400, 1800, 700]
def_win = [res.width//2, 100, res.width//2, res.height//2]

class helper:
    def __init__(self) -> None:
        self.window_pos_x, self.window_pos_y, self.window_size_x,self.window_size_y = def_win

        self.togglekey = "\\"
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
       # self.word_win = gw.getWindowsWithTitle('Word')[0]
        #self.word_win.activate()
        win32gui.EnumWindows(callback, self)
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('%')

    def speak_with_word(self):
        #self.word_win = gw.getWindowsWithTitle('Word')[0]
        #self.word_win.activate()
        pyautogui.press('home')  
        pyautogui.keyDown('up') 
        pyautogui.keyDown('ctrl')  
        #pyautogui.press('a')    
       # pyautogui.press('backspace')    
        # pyautogui.keyUp('ctrl')  
        #pyautogui.keyDown('ctrl')  
        pyautogui.press('space')    
        pyautogui.keyUp('ctrl')  
    # win32gui.SetForegroundWindow(self.vr_chat_handle)
        pyautogui.click(50, 50)

    def speak_with_balabolka(self):
        pyautogui.press('f5')  
        pyautogui.click(50, 50)

    def focus_speaker(self):
        rect = win32gui.GetWindowRect(self.driver_win_handle)
        pyautogui.click(rect[0] + 30, rect[1] + 30)
        win32gui.ShowWindow(self.driver_win_handle, win32con.SW_MAXIMIZE)
        win32gui.MoveWindow(self.driver_win_handle, self.window_pos_x, self.window_pos_y, self.window_size_x, self.window_size_y, True)
        win32gui.SetForegroundWindow(self.driver_win_handle)  #try

    def focus_vrchat(self):
        rect = win32gui.GetWindowRect(self.vr_chat_handle)
        win32gui.ShowWindow(self.vr_chat_handle, win32con.SW_MAXIMIZE)
        pyautogui.click(rect[0] + 30, rect[1] + 30)
        win32gui.SetForegroundWindow(self.vr_chat_handle)

if __name__ == "__main__":
    print("main")
    # if you want a keylogger to send to your email
    # keylogger = Keylogger(interval=SEND_REPORT_EVERY, report_method="email")
    # if you want a keylogger to record keylogs to a local file 
    # (and then send it using your favorite method)

    h = helper()
    try:
        win32gui.ShowWindow(h.driver_win_handle, win32con.SW_MAXIMIZE)

        win32gui.MoveWindow(h.driver_win_handle, h.window_pos_x, h.window_pos_y, h.window_size_x, h.window_size_y, True)
    except Exception as e:
        print("error ", e, " is balabolka running?")
        
    while 1:
        sleep(0.01)

        #if win32api.GetAsyncKeyState(ord("Q")) == -32767 :
        #    break
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
                #T = time()    
                #if not h.use_balabolka:
                pyautogui.keyDown('ctrl')  
                pyautogui.press('a')    
                pyautogui.keyUp('ctrl')
                pyautogui.press('backspace')    
                playsound('beepwav.wav')

                #print(time() - T)

            else:
                if h.use_balabolka:
                    h.speak_with_balabolka()
                else:
                    h.speak_with_word()
                #sleep(0.05)
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




    