from pywinauto import application
import time
from get_passcode import get_passcode
from tkinter.messagebox import showerror
import tkinter as tk
from tkinter import *
import os
import threading
import pickle

main_app = 0
main_window = 0
global_username = ""
global_password = ""
global_pin = ""
global_cisco_path = ""
global_passcode_path = ""

Error_Printing = False

class Connection_Thread (threading.Thread):
    def __init__(self, threadID, name):
        threading.Thread.__init__(self)
        self.threadID = threadID
        self.name = name
    def run(self):
        global global_username
        
        connect_to_vpn(global_username, global_password, global_pin, global_cisco_path, global_passcode_path)
        
        # Clear global username again to activate connect button
        global_username = ""
        main_window.labelText.set("Connected successfully!")
        main_window.status_label.config(fg='#50C878', font='Helvetica 12 bold')
        

def connect_to_vpn(username, password, PIN, cisco_path, passcode_path):

    app = application.Application(backend="uia")
    
    # Check if cisco appication is already running and kill it for now if so
    # Note: If not running, app.connect() will raise exception
    try:
        app.connect(path=cisco_path)
        app.kill(soft=False)
    except Exception as e:
        if Error_Printing:
            print(e)
    
    # Start application again
    app.start(cisco_path)

    while not app.window_():
        time.sleep(.5)

    # Get top window and get its title
    wnd = None
    wnd_up = False
    while not wnd_up:
        try:
            wnd = app.top_window()
            wnd_up = True
        except Exception as e:
            if Error_Printing:
                print(e)
            try:
                app.connect(title="Cisco AnyConnect Secure Mobility Client", found_index=0)
            except Exception as e:
                #Show "Already connected to VPN!" message
                main_window.event_generate("<<error_event>>", when="tail", state=4) # trigger event in main thread
                return
            wnd_up = False
    
    if Error_Printing:
        print(wnd.texts()[0])
    
    if wnd.window(title="Disconnect", control_type="Button").exists():
        #popup_showError("Already connected to VPN!")
        main_window.event_generate("<<error_event>>", when="tail", state=4) # trigger event in main thread
        return

    wnd.Connect.click()

    # Wait for username password window to appear
    while not wnd.window(title_re='Cisco AnyConnect | * - 2FA').exists():
        time.sleep(.5)

    # Get reference to username and password window and set username and password and click 'OK'
    uname_pwd_win = wnd.window(title_re='Cisco AnyConnect | * - 2FA')

    uname_pwd_win.set_focus()

    uname_field = uname_pwd_win.window(title="Username:", control_type="Edit")

    uname_field.set_text(username)

    pwd_field = uname_pwd_win.window(title="Password:", control_type="Edit")

    pwd_field.set_text(password)

    uname_pwd_win.OK.click()

    # Wait for OTP window to appear
    while not wnd.window(title_re='Cisco AnyConnect | * - 2FA').exists():
        time.sleep(.5)

    # Get passcode (OTP) from passcode app
    passcode = get_passcode(PIN)

    # Enter OTP and click continue to same window
    otp_field = uname_pwd_win.window(title="Answer:", control_type="Edit")
    otp_field.set_text(passcode)
    uname_pwd_win.Continue.click()

    # Wait for Accept dialog window to appear
    while not wnd.window(title="Cisco AnyConnect", control_type="Window").exists():
        time.sleep(.5)

    # Get accept dialog window, set focus and click 'Accept'
    accept_window = wnd.window(title="Cisco AnyConnect", control_type="Window")
    accept_window.set_focus()
    accept_window.Accept.click()
    
def popup_showError(error):
    showerror("Error", error)
        
def eventhandler(evt):
    if (evt.state == 4):
        popup_showError("Already connected to VPN!")
        
def enter_key_handler(evt):
    main_window.connect_btn.invoke()
 
def main():

    global main_app
    
    # Create main window
    class MainApp(tk.Tk):
        def __init__(self):
            tk.Tk.__init__(self)
            self1 = self
            self.title('VPN Auto Connect')
            
            #self.state('zoomed')
            self.minsize(1000, 500)
            
            global main_window
            main_window = self
            
            # Add username, password and PIN fields
            padding = 10
            
            self.paths_frame = Frame(self)
            self.paths_frame.grid(row=0, padx=(10, 10))
            self.credentials_frame = Frame(self)
            self.credentials_frame.grid(row=1, padx=(10, 10))
            self.labelText = tk.StringVar(self)
            self.status_label = tk.Label(self, textvariable=self.labelText)
            self.status_label.grid(row=2, pady=(10, 10))
            
            tk.Label(self.paths_frame, text="Path to Cisco AnyConnect App:").grid(row=0, pady=(padding, 0))
            tk.Label(self.paths_frame, text="Path to Passcode App").grid(row=2, pady=(padding, 0))
            
            self.paths_frame.cisco_path_entry = tk.Entry(self.paths_frame, width=75)
            self.paths_frame.passcode_path_entry = tk.Entry(self.paths_frame, width=75)
            self.paths_frame.cisco_path_entry.grid(row=1, pady=(0, padding), sticky = E+W)
            self.paths_frame.passcode_path_entry.grid(row=3, pady=(0, padding), sticky = E+W)
            
            tk.Label(self.credentials_frame, text="MGC Username:").grid(row=0, pady=(padding, padding))
            tk.Label(self.credentials_frame, text="MGC Password").grid(row=1, pady=(padding, padding))
            tk.Label(self.credentials_frame, text="Passcode App PIN").grid(row=2, pady=(padding, padding))
            
            self.credentials_frame.username_entry = tk.Entry(self.credentials_frame)
            self.credentials_frame.password_entry = tk.Entry(self.credentials_frame, show="*")
            self.credentials_frame.pin_entry = tk.Entry(self.credentials_frame, show="*")
            
            self.credentials_frame.username_entry.grid(row=0, column=1, pady=(padding, padding))
            self.credentials_frame.password_entry.grid(row=1, column=1, pady=(padding, padding))
            self.credentials_frame.pin_entry.grid(row=2, column=1, pady=(padding, padding))
            
            # Load default app paths
            cisco_path = r"C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe"
            passcode_path = r"C:\Program Files (x86)\Passcode\Passcode.exe"
            
            if os.path.exists('credentials.txt'):
                with open("credentials.txt","rb") as handle:
                    credentials = pickle.loads(handle.read())
                    
                    if os.path.exists(credentials['cisco_path']):
                        self.paths_frame.cisco_path_entry.insert(0, credentials['cisco_path'])
                        
                    elif os.path.exists(cisco_path):
                        self.paths_frame.cisco_path_entry.insert(0, cisco_path)
                        
                    if os.path.exists(credentials['passcode_path']):
                        self.paths_frame.passcode_path_entry.insert(0, credentials['passcode_path'])
                        
                    elif os.path.exists(passcode_path):
                        self.paths_frame.passcode_path_entry.insert(0, passcode_path)
                        
                    self.credentials_frame.username_entry.insert(0, credentials['username'])
                    self.credentials_frame.password_entry.insert(0, credentials['password'])
                    self.credentials_frame.pin_entry.insert(0, credentials['pin'])
                    
            else:
                
                if os.path.exists(cisco_path):
                    self.paths_frame.cisco_path_entry.insert(0, cisco_path)
                    
                if os.path.exists(passcode_path):
                    self.paths_frame.passcode_path_entry.insert(0, passcode_path)
            
            self.check_var = tk.IntVar(self)
            self.check_var.set(1)
            self.save_cred_checkbox = Checkbutton(self.credentials_frame, text="Save Credentials?", 
                                                  variable=self.check_var)
            self.save_cred_checkbox.grid(row=3, column=0)
            
            # Add Connect button
            self.connect_btn = tk.Button(self.credentials_frame, text = 'Connect', bg = 'white', fg = 'black', 
                                    command= lambda: self.connect_vpn(self.credentials_frame.username_entry.get(), 
                                                                      self.credentials_frame.password_entry.get(), 
                                                                      self.credentials_frame.pin_entry.get(), 
                                                                      self.paths_frame.cisco_path_entry.get(), 
                                                                      self.paths_frame.passcode_path_entry.get(), 
                                                                      bool(self.check_var.get()))) # bg = 'seashell2'
            self.connect_btn.grid(row=3, column=1)
            
            # Add status bar
            self.status_bar = tk.Label(self, text="From Waseem Abbas 🙂", bd=1, relief=SUNKEN, anchor=W)
            self.status_bar.grid(row=3, pady=(10, 0), sticky = E+W)
            
            #self.grid_columnconfigure(0, weight = 1)
            #self.grid_columnconfigure(1, weight = 1)

            self.grid_rowconfigure(1, weight = 1)
            self.grid_rowconfigure(3, weight = 1)
            
            # Set event handler and Enter key press handler
            self.bind("<<error_event>>", eventhandler)
            self.bind('<Return>', enter_key_handler)
            
            
        def connect_vpn(self, username, password, pin, cisco_path, passcode_path, check_var): #username_entry, password_entry, pin_entry):
            
            global global_username
            global global_password
            global global_pin
            global global_cisco_path
            global global_passcode_path
            
            # Strip double quotes from paths if any
            cisco_path = cisco_path.strip('"')
            passcode_path = passcode_path.strip('"')
            
            if len(global_username) != 0:
                popup_showError("Connection is already in progress!")
                return
            
            if len(cisco_path) == 0:
                popup_showError("Cisco app path is empty!");
                return
            
            elif not os.path.exists(cisco_path):
                popup_showError("Invalid Cisco app path!")
                return
                
            if len(passcode_path) == 0:
                popup_showError("Passcode app path is empty!");
                return
            
            elif not os.path.exists(passcode_path):
                popup_showError("Invalid Passcode app path!")
                return
                
            if ((len(username) == 0) or (len(password) == 0) or (len(pin) == 0)):
                popup_showError("Please enter all credentials!")
                return
                
            elif len(pin) != 4:
                popup_showError("Passcode App's PIN must consist of 4 digits!")
                return
            
            # Local copy of credentials
            username_local = username
            password_local  = password
            pin_local = pin
            
            if not check_var:
                username_local = ""
                password_local = ""
                pin_local = ""
            
            credentials = {'username': username_local, 
                            'password': password_local,
                            'pin': pin_local,
                            'cisco_path': cisco_path,
                            'passcode_path': passcode_path}
                            
            with open("credentials.txt","wb+") as handle:
                pickle.dump(credentials, handle)
            
            #label = tk.Label(self, text="Connecting to VPN....")
            #label.grid(row=2, pady=(10, 10))
            
            #self.status_label.config(text= "Connecting to VPN....")
            #self.status_label.Invalidate
            
            self.labelText.set("Connecting to VPN....")
            self.status_label.config(fg='#000000')
            
            global_username = username
            global_password = password
            global_pin = pin
            global_cisco_path = cisco_path
            global_passcode_path = passcode_path
            
            # Create new thread
            vpnConnectThread = Connection_Thread(1, "ConnectionThread")

            # Start new thread as a deamon
            vpnConnectThread.setDaemon(True)
            vpnConnectThread.start()
            
            #connect_to_vpn(username, password, pin, cisco_path, passcode_path)
            
            #self.status_label.config(text= "Connected successfully!")
            #self.labelText.set("Connected successfully!")

    main_app = MainApp()
    main_app.mainloop()
    

if __name__ == '__main__':

    root = Tk()
    root.withdraw()
    main()
    os._exit(0)
    
    