from pywinauto import application
import time
import win32clipboard

Error_Printing = False

#########################################################################
#            Launch Passcode App, enter PIN and return passcode         #
#########################################################################
def get_passcode(pin):
    
    passcode_program_path = r"C:\Program Files (x86)\Passcode\Passcode.exe"
    passcode_app = application.Application(backend="uia")
    passcode_app.start(passcode_program_path)
    
    while not passcode_app.window_():
        time.sleep(.5)
        
    if Error_Printing:
        print ("started passcode app")

    # Wait for PIN dialog window to appear
    # time.sleep(7)

    # Connect to first index window of this app and get its top window
    passcode_app_connect = application.Application(backend="uia")
    
    Connected = False
    while not Connected:
        try:
            passcode_app_connect.connect(title="Passcode", control_type="Window", found_index=0)
            Connected = True
        except Exception as e:
            if Error_Printing:
                print(e)
            Connected = False
            
    Connected = False

    passcode_top_window = passcode_app_connect.top_window()

    passBox1 = passcode_top_window.window(auto_id="passBox1", control_type="Edit")
    
    # This while is due to the fact thta Passocde app consists of multiple windows that appear at startup 
    # like splash etc. This loop ensures that we wait for the actual window that accepts login PIN
    passBox1_update_status = False
    while not passBox1_update_status:
        try:
            passBox1.set_text(pin[0])
            passBox1_update_status = True
        except Exception as e:
            if Error_Printing:
                print(e)
            # Connect to first index window of this app again, get its top window 
            # and try to get the passBox1 field again
            passcode_app_connect = application.Application(backend="uia")
            passcode_app_connect.connect(title="Passcode", control_type="Window", found_index=0)
            passcode_top_window = passcode_app_connect.top_window()
            passBox1 = passcode_top_window.window(auto_id="passBox1", control_type="Edit")
            passBox1_update_status = False

    passBox2 = passcode_top_window.window(auto_id="passBox2", control_type="Edit")
    passBox2.set_text(pin[1])

    passBox3 = passcode_top_window.window(auto_id="passBox3", control_type="Edit")
    passBox3.set_text(pin[2])

    passBox4 = passcode_top_window.window(auto_id="passBox4", control_type="Edit")
    passBox4.set_text(pin[3])

    # Now that PIN has be populated, click 'Enter' button to get Passcode
    enter_button = passcode_top_window.window(title="Enter", auto_id="enterBtn", control_type="Button")
    enter_button.click()
    
    # Put some delay so that passcode windows opens up properly
    time.sleep(1)

    # Now connect to second instance window of this app (the one that shows passcode) and get its top window
    passcode_app_connect.connect(title="Passcode", control_type="Window", found_index=0)

    passcode_top_window = passcode_app_connect.top_window()
    
    # passcode_top_window.print_control_identifiers();
    #return
    
    otp_field = passcode_top_window.window(auto_id="otpLabel", control_type="Text")
    passcode = otp_field.element_info.name
    if Error_Printing:
        print (passcode)
    
    # Close Passcode app
    passcode_top_window.close()
    
    return passcode
    