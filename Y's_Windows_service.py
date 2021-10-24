
import win32service # This module implements the Win32 service functions
import win32serviceutil # This module provides some handy utilities that utilize the raw API
import win32event
import datetime
import os,sys


# /////////////////////////////////////////////////////////////  1/4  ////////////////////////////////////////////////////////////////


# ____ Creating the Windows service in Python ____

class myWinService(win32serviceutil.ServiceFramework):

# 1. Creating a log file ___________________

    logfile = r'C:\\MY_LOGS\\myWinService.log'
    f = open(logfile, 'w+')

# 2. Naming the service ___________________

    _svc_name_ = 'my_Win_Service'   # Type 'net start myWinService' / 'net stop myWinService' in the command line to start/stop this service

    _svc_display_name_ = 'My cool windows service' # This text shows up as the service name in the Service Control Manager ('SCM')

    _svc_description_ = 'RThis service runs my script automatically' # this text shows up as the description in the SCM



# 3. Initial functions for the service ___________________

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self,args)
        # create an event to listen for stop requests on
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)


# /////////////////////////////////////////////////////////////  2/4  ////////////////////////////////////////////////////////////////



    def SvcStop(self): # Stop the windows service
        # tell the SCM we're shutting down
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)

        # fire the stop event
        win32event.SetEvent(self.hWaitStop)


# /////////////////////////////////////////////////////////////

    # When service is activated, run this script:
    def SvcDoRun(self):

        Script_name = 'myScript.py'
        Python_exe_path = r'C:\\Python\\Python37-32\\python.exe'

        rc = None

        self.timeout=60000

        #___________________________________________________________________________________________________________

        # while service is running (no stop event fired):
        while rc != win32event.WAIT_OBJECT_0:

            try:
                # Run the 'myScript.py' script
                os.system(Python_exe_path + " " + os.path.join(os.path.dirname(os.path.abspath(__file__)),Script_name))

            except:
                self.f.write("Can't run your script, sir!")

        #___________________________________________________________________________________________________________

            # block for tot seconds and listen for a stop event
            rc = win32event.WaitForSingleObject(self.hWaitStop, self.timeout)

        try:
            self.f.write(str(datetime.datetime.now()) + "     'my_Win_Service' Windows service stopped!" + '\n\n')
            self.f.close()


        except:
            self.f.write(str(datetime.datetime.now()) + "     Can't stop 'my_Win_Service' Windows service!" + '\n\n')
            self.f.close()



# /////////////////////////////////////////////////////////////  3/4  ////////////////////////////////////////////////////////////////


# 4. Launching the service  ___________________

if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(my_Win_Service)

