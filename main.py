from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
import time
from win32gui import GetForegroundWindow
import win32process
import wmi
import keyboard


def key_change_vol(value, exe):
    try:
        new_val = change_vol(value, exe)
    except TypeError:
        pass
    else:
        print(f'{exe}: {round(new_val, 2)}')
    time.sleep(0.2)


def main():
    keypress = False
    exe = get_current_process()
    old_exe = None

    while True:
        exe = get_current_process()

        if keyboard.is_pressed('ctrl+shift+q'):
            key_change_vol(0.1, exe)
        elif keyboard.is_pressed('ctrl+shift+a'):
            key_change_vol(-0.1, exe)

        if old_exe != exe:
            print(f'Switched to {exe}: {get_vol(exe)}')
            old_exe = exe
        time.sleep(0.1)


def change_vol(value, process):
    cur_val = get_vol(process)
    cur_val = cur_val + value
    cur_val = min(1.0, max(0.0, cur_val))
    ret_val = cur_val
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        if session.Process and session.Process.name() == process:
            volume.SetMasterVolume(cur_val, None)
    return cur_val


def get_vol(process):
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        if session.Process and session.Process.name() == process:
            return volume.GetMasterVolume()


def get_current_process():
    c = wmi.WMI()
    exe = None
    try:
        _, pid = win32process.GetWindowThreadProcessId(GetForegroundWindow())
        for p in c.query('SELECT Name FROM Win32_Process WHERE ProcessId = %s' % str(pid)):
            exe = p.Name
            break
    except:
        return None
    else:
        return exe


if __name__ == '__main__':
    main()

