import time
import P4Summary
import CopyBuildBackup
import JiraReadme
import os

timefreq = 1800

def main():
    while 1:

        print('')
        print(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
        start_time = time.time()
        try:
            os.system('E:\\Tool\\Install_Opengrok\\syn_opengrok.bat')
            CopyBuildBackup.main()
            print('')
            P4Summary.main()
            JiraReadme.main()

        except Exception as e:
            # log(str(e), level.error)
            print(str(e))
            pass
        finally:
            print(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
            print("time elapsed: {:.2f}s".format(time.time() - start_time))
            print('Idle..............................')
            time.sleep(timefreq)

if __name__ == "__main__":
    main()
