import const
import common

import os
from ftplib import FTP

import imaplib
import time

import socket
from smb.SMBConnection import SMBConnection

def main():
        ### get mail
        try:
            url = const.trend_mail_server
            mail_conn = imaplib.IMAP4_SSL(url, 993)
            user, password = (const.trend_account, const.trend_password)
            mail_conn.login(user, password)
            mail_conn.select('INBOX/Other/Build')
            print('Get Mail from Trend...')

            results, data = mail_conn.search(None, 'UnSeen')
            # results, data = conn.search(None, 'All')
            msg_ids = data[0]
            msg_id_list = msg_ids.split()

            if len(msg_id_list) > 0:
                latest_email_id = msg_id_list[-1]
                result, data = mail_conn.fetch(latest_email_id, "(RFC822)")
                raw_email = data[0][1]

                raw_email = str(raw_email, 'utf-8')
                from email.parser import Parser
                p = Parser()
                msg = p.parsestr(raw_email)

                msg_contant = process_multipart_message(msg)
                if msg_contant.upper().find('BACKUP') > 0 and  msg_contant.find('Pass') > 0:

                    build_number = msg_contant.split('TMCM-SP3_6.0_6.0.0.')[1].split(
                        '                                                              <br>')[0]
                    lang = msg_contant.split('Build Language	:	')[1].split('                     <br>')[0]
                    # path = msg_contant.split('FTP Path	: ')[1].split(' </div> ')[0]

                    try:
                        print('Get Build from FTP...' + build_number + '  ' + lang)
                        ### download build
                        download_prepack_file(build_number, lang)

                        print('Copy File to NAS...' + build_number + '  ' + lang)
                        ### save file to nas
                        save_file_to_smb(build_number, lang)
                        results, data = mail_conn.store(latest_email_id, '+FLAGS', '\\Seen')

                        print('All Done...' + build_number + '  ' + lang)

                    except Exception as e:
                        results, data = mail_conn.store(latest_email_id, '-FLAGS', '\\Seen')
                        print(str(e))
                        # log(str(e), level.error)
                        pass
                else:
                    results, data = mail_conn.store(latest_email_id, '+FLAGS', '\\Seen')

        except Exception as e:
            # log(str(e), level.error)
            print(str(e))
            pass


def process_multipart_message(message):
    rtn = ''
    if message.is_multipart():
        for m in message.get_payload():
            rtn += process_multipart_message(m)
    else:
        rtn += message.get_payload()
    return rtn


def download_prepack_file(build_number, language):
    build_dst = const.HF_Working_Folder + "\\build\\" + language + "\\B" + build_number
    filename = 'Prepack.zip'

    if not os.path.exists(build_dst + "\\" + filename):
        if not os.path.exists(build_dst):
            build_dst = common.create_folder(build_dst)
        prepack_download_from_ftp(filename, build_dst, build_number, language)


def prepack_download_from_ftp(filename, download_folder, build_number, language):
    build_src_pkg_path = get_prepack_src_file_path(build_number, language)
    # log("Download {0} from {1}".format(filename, build_src_pkg_path), level.info)
    print("Download {0} from {1}".format(filename, build_src_pkg_path))
    try:
        os.chdir(download_folder)
        ftp = FTP(const.build_ftp_server)
        ftp.login(const.build_ftp_user, const.build_ftp_pwd)
        ftp.cwd(build_src_pkg_path)
        fhandle = open(filename, 'wb')
        ftp.retrbinary('RETR ' + filename, fhandle.write)
        fhandle.close()
    except:
        os.chdir(const.HF_Working_Folder + '\Build')
        raise
    # log("Download {0} finished.".format(filename), level.info)
    print("Download {0} finished.".format(filename))


def get_prepack_src_file_path(build_number, language):
    if language == 'jp':
        return "Build/TMCM-SP3/6.0/win32/{0}/Rel/{1}/release/output/".format('ja', build_number)
    elif language == 'cn':
        return "Build/TMCM-SP3/6.0/win32/{0}/Rel/{1}/release/output/".format('zh_CN', build_number)
    else:
        return "Build/TMCM-SP3/6.0/win32/{0}/Rel/{1}/release/output/".format(language, build_number)


def save_file_to_smb(build_number, language):
    try:
        name = socket.gethostbyaddr(const.nas_ip)
        ipGet = socket.gethostbyname(name[0])
        print(name, ipGet, sep='\n')

        remote_name = name[0]
        smb_conn = SMBConnection(const.nas_account, const.nas_password, 'any_name', remote_name)
        assert smb_conn.connect(const.nas_ip, timeout=30)

        prepack_path = const.HF_Working_Folder + '\\build\\' + language
        f = open(prepack_path + '\\B' + build_number + '\\Prepack.zip', 'rb')

        if language == 'ja':
            language = 'JP'
        elif language == 'zh_CN':
            language = 'SC'

        smb_conn.storeFile('Backup_Build', '\\' + language + '\\Prepack_' + build_number + '.zip', f)
        smb_conn.close()
        f.close()

        print('Remove local file...')
        os.remove(prepack_path + '\\B' + build_number + '\\Prepack.zip')

    except Exception as e:
        # log(str(e), level.error)
        print(str(e))
        raise Exception('Failed! Save file to NAS unsuccessfully.')

if __name__ == '__main__':
    main()
