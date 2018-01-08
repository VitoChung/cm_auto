import os

def create_folder(_folder):
    if os.path.isdir(_folder):
        run_console_cmd("rmdir /S /Q " + _folder)
        os.makedirs(_folder)
    else:
        os.makedirs(_folder)
    return _folder

def run_console_cmd(cmd):
    print (os.popen("echo " + cmd).read())
    print (os.popen(cmd).read())
