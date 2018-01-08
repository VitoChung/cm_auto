import time
import P4Summary
import CopyBuildBackup
import os

timefreq = 1800

def main():
    while 1:

        print('')
        print(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))

        try:
            os.system('E:\\Tool\\Install_Opengrok\\syn_opengrok.bat')
            CopyBuildBackup.main()
            print('')
            P4Summary.main()

        except Exception as e:
            # log(str(e), level.error)
            print(str(e))
            pass
        finally:
            print('Idle..............................')
            time.sleep(timefreq)

if __name__ == "__main__":
    main()
