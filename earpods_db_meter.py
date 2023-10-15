import os, errno
import pyaudio
from scipy.signal import lfilter
import numpy
from tkinter import *
from tkinter.ttk import *
from tk_tools import *
from tkinter import messagebox
import ctypes
import comtypes
from ctypes import wintypes
from pynput.keyboard import Key, Controller
from time import sleep
from idlelib.tooltip import Hovertip
from win10toast import ToastNotifier
from threading import Thread
import time
import sys
def get_resource_path(relative_path):
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)
try:
    from winreg import *
except:
    reg_present=False
    messagebox.askokcancel('Limited Features', "Registry not present. Dosimeter disabled. OK to continue, Cancel to quit.", icon='warning')
else:
    reg_present=True
if reg_present:
    CreateKeyEx(OpenKey(HKEY_CURRENT_USER, 'Software', reserved=0, access=KEY_ALL_ACCESS), 'sserver\Decibel Meter', reserved=0)
    try:
        dosi_enabled=bool(QueryValueEx(OpenKey(OpenKey(OpenKey(HKEY_CURRENT_USER, 'Software', reserved=0, access=KEY_ALL_ACCESS), 'sserver', reserved=0, access=KEY_ALL_ACCESS), 'Decibel Meter', reserved=0, access=KEY_ALL_ACCESS), 'DosimeterEnabled')[0])
    except OSError:
        SetValueEx(OpenKey(OpenKey(OpenKey(HKEY_CURRENT_USER, 'Software', reserved=0, access=KEY_ALL_ACCESS), 'sserver', reserved=0, access=KEY_ALL_ACCESS), 'Decibel Meter', reserved=0, access=KEY_ALL_ACCESS), 'DosimeterEnabled', 0, REG_DWORD, 0)
        dosi_enabled=False
        messagebox.showwarning('Registry Error', 'Error reading settings. Resetting to default...')
else:
    dosi_enabled=False

def toggle_dosi():
    global dosi_enabled
    dosi_enabled=enable_dosi.instate(['selected'])
MMDeviceApiLib = comtypes.GUID(
    '{2FDAAFA3-7523-4F66-9957-9D5E7FE698F6}')
IID_IMMDevice = comtypes.GUID(
    '{D666063F-1587-4E43-81F1-B948E807363F}')
IID_IMMDeviceCollection = comtypes.GUID(
    '{0BD7A1BE-7A1A-44DB-8397-CC5392387B5E}')
IID_IMMDeviceEnumerator = comtypes.GUID(
    '{A95664D2-9614-4F35-A746-DE8DB63617E6}')
IID_IAudioEndpointVolume = comtypes.GUID(
    '{5CDF2C82-841E-4546-9722-0CF74078229A}')
CLSID_MMDeviceEnumerator = comtypes.GUID(
    '{BCDE0395-E52F-467C-8E3D-C4579291692E}')

# EDataFlow
eRender = 0 # audio rendering stream
eCapture = 1 # audio capture stream
eAll = 2 # audio rendering or capture stream
keyboard=Controller()
# ERole
eConsole = 0 # games, system sounds, and voice commands
eMultimedia = 1 # music, movies, narration
eCommunications = 2 # voice communications

LPCGUID = REFIID = ctypes.POINTER(comtypes.GUID)
LPFLOAT = ctypes.POINTER(ctypes.c_float)
LPDWORD = ctypes.POINTER(wintypes.DWORD)
LPUINT = ctypes.POINTER(wintypes.UINT)
LPBOOL = ctypes.POINTER(wintypes.BOOL)
PIUnknown = ctypes.POINTER(comtypes.IUnknown)
CHUNKS = [4096, 9600]
CHUNK = CHUNKS[1]
toaster = ToastNotifier()
FORMAT = pyaudio.paInt16
CHANNEL = 1 
RATES = [44300, 48000]
RATE = RATES[1]
class IMMDevice(comtypes.IUnknown):
    _iid_ = IID_IMMDevice
    _methods_ = (
        comtypes.COMMETHOD([], ctypes.HRESULT, 'Activate',
            (['in'], REFIID, 'iid'),
            (['in'], wintypes.DWORD, 'dwClsCtx'),
            (['in'], LPDWORD, 'pActivationParams', None),
            (['out','retval'], ctypes.POINTER(PIUnknown), 'ppInterface')),
        comtypes.STDMETHOD(ctypes.HRESULT, 'OpenPropertyStore', []),
        comtypes.STDMETHOD(ctypes.HRESULT, 'GetId', []),
        comtypes.STDMETHOD(ctypes.HRESULT, 'GetState', []))

PIMMDevice = ctypes.POINTER(IMMDevice)

class IMMDeviceCollection(comtypes.IUnknown):
    _iid_ = IID_IMMDeviceCollection

PIMMDeviceCollection = ctypes.POINTER(IMMDeviceCollection)

class IMMDeviceEnumerator(comtypes.IUnknown):
    _iid_ = IID_IMMDeviceEnumerator
    _methods_ = (
        comtypes.COMMETHOD([], ctypes.HRESULT, 'EnumAudioEndpoints',
            (['in'], wintypes.DWORD, 'dataFlow'),
            (['in'], wintypes.DWORD, 'dwStateMask'),
            (['out','retval'], ctypes.POINTER(PIMMDeviceCollection),
             'ppDevices')),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'GetDefaultAudioEndpoint',
            (['in'], wintypes.DWORD, 'dataFlow'),
            (['in'], wintypes.DWORD, 'role'),
            (['out','retval'], ctypes.POINTER(PIMMDevice), 'ppDevices')))
    @classmethod
    def get_default(cls, dataFlow, role):
        enumerator = comtypes.CoCreateInstance(
            CLSID_MMDeviceEnumerator, cls, comtypes.CLSCTX_INPROC_SERVER)
        return enumerator.GetDefaultAudioEndpoint(dataFlow, role)

class IAudioEndpointVolume(comtypes.IUnknown):
    _iid_ = IID_IAudioEndpointVolume
    _methods_ = (
        comtypes.STDMETHOD(ctypes.HRESULT, 'RegisterControlChangeNotify', []),
        comtypes.STDMETHOD(ctypes.HRESULT, 'UnregisterControlChangeNotify', []),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'GetChannelCount',
            (['out', 'retval'], LPUINT, 'pnChannelCount')),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'SetMasterVolumeLevel',
            (['in'], ctypes.c_float, 'fLevelDB'),
            (['in'], LPCGUID, 'pguidEventContext', None)),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'SetMasterVolumeLevelScalar',
            (['in'], ctypes.c_float, 'fLevel'),
            (['in'], LPCGUID, 'pguidEventContext', None)),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'GetMasterVolumeLevel',
            (['out','retval'], LPFLOAT, 'pfLevelDB')),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'GetMasterVolumeLevelScalar',
            (['out','retval'], LPFLOAT, 'pfLevel')),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'SetChannelVolumeLevel',
            (['in'], wintypes.UINT, 'nChannel'),
            (['in'], ctypes.c_float, 'fLevelDB'),
            (['in'], LPCGUID, 'pguidEventContext', None)),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'SetChannelVolumeLevelScalar',
            (['in'], wintypes.UINT, 'nChannel'),
            (['in'], ctypes.c_float, 'fLevel'),
            (['in'], LPCGUID, 'pguidEventContext', None)),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'GetChannelVolumeLevel',
            (['in'], wintypes.UINT, 'nChannel'),
            (['out','retval'], LPFLOAT, 'pfLevelDB')),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'GetChannelVolumeLevelScalar',
            (['in'], wintypes.UINT, 'nChannel'),
            (['out','retval'], LPFLOAT, 'pfLevel')),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'SetMute',
            (['in'], wintypes.BOOL, 'bMute'),
            (['in'], LPCGUID, 'pguidEventContext', None)),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'GetMute',
            (['out','retval'], LPBOOL, 'pbMute')),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'GetVolumeStepInfo',
            (['out','retval'], LPUINT, 'pnStep'),
            (['out','retval'], LPUINT, 'pnStepCount')),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'VolumeStepUp',
            (['in'], LPCGUID, 'pguidEventContext', None)),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'VolumeStepDown',
            (['in'], LPCGUID, 'pguidEventContext', None)),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'QueryHardwareSupport',
            (['out','retval'], LPDWORD, 'pdwHardwareSupportMask')),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'GetVolumeRange',
            (['out','retval'], LPFLOAT, 'pfLevelMinDB'),
            (['out','retval'], LPFLOAT, 'pfLevelMaxDB'),
            (['out','retval'], LPFLOAT, 'pfVolumeIncrementDB')))
    @classmethod
    def get_default(cls):
        endpoint = IMMDeviceEnumerator.get_default(eRender, eMultimedia)
        interface = endpoint.Activate(cls._iid_, comtypes.CLSCTX_INPROC_SERVER)
        return ctypes.cast(interface, ctypes.POINTER(cls))
from scipy.signal import bilinear
def close():
    global appclosed
    win.destroy()
    recording=False
    appclosed=True
    if dosi_enabled:
        SetValueEx(OpenKey(OpenKey(OpenKey(HKEY_CURRENT_USER, 'Software', reserved=0, access=KEY_ALL_ACCESS), 'sserver', reserved=0, access=KEY_ALL_ACCESS), 'Decibel Meter', reserved=0, access=KEY_ALL_ACCESS), 'DosimeterEnabled', 0, REG_DWORD, 1)
    else:
        SetValueEx(OpenKey(OpenKey(OpenKey(HKEY_CURRENT_USER, 'Software', reserved=0, access=KEY_ALL_ACCESS), 'sserver', reserved=0, access=KEY_ALL_ACCESS), 'Decibel Meter', reserved=0, access=KEY_ALL_ACCESS), 'DosimeterEnabled', 0, REG_DWORD, 0)
    stream.stop_stream()
    stream.close()
    pa.terminate()
    sys.exit()
def A_weighting(fs):
    f1 = 20.598997
    f2 = 107.65265
    f3 = 737.86223
    f4 = 12194.217
    A1000 = 1.9997

    NUMs = [(2*numpy.pi * f4)**2 * (10**(A1000/20)), 0, 0, 0, 0]
    DENs = numpy.polymul([1, 4*numpy.pi * f4, (2*numpy.pi * f4)**2],
                   [1, 4*numpy.pi * f1, (2*numpy.pi * f1)**2])
    DENs = numpy.polymul(numpy.polymul(DENs, [1, 2*numpy.pi * f3]),
                                 [1, 2*numpy.pi * f2])
    return bilinear(NUMs, DENs, fs)
def rms_flat(a):
    return numpy.sqrt(numpy.mean(numpy.absolute(a)**2))
pa = pyaudio.PyAudio()
n=1
f=None
recording=False
while True:
    try:
        if pa.get_device_info_by_index(n)["name"]=='Hi-Fi Cable Output (VB-Audio Hi-Fi Cable)':
            stream = pa.open(format = FORMAT,
                channels = CHANNEL,
                rate = RATE,
                input = True,
                frames_per_buffer = CHUNK,)
            try:
                os.system('SoundVolumeView.exe /SetAppDefault "VB-Audio Hi-Fi Cable\Device\CABLE Output\Capture" 1 '+str(os.getpid()))
            except OSError:
                messagebox.showwarning('Loopback Device Found', 'Setting VB-Cable failed. Select the VB Cable by going to Settings>Sound>Volume mixer>Decibel Meter and select CABLE Output as input device')
            break
        else:
            n+=1
    except OSError:
        messagebox.showerror('No Loopback Device found', 'There is no Virtual Cable installed on this machine. This app will now close.')
        appclosed=True
        os._exit(1)
def update_limiter():
    if sb.get()=='Off':
        root.limiter=194
    else:
        root.limiter=int(sb.get()[:3])
def reset():
    global start, dosimeter_times, runTime, x, y, plot1, gr_start
    dosimeter_times={'82dB':0, '85dB':0, '88dB':0, '91dB':0, '94dB':0, '97dB':0, '100dB':0, '103dB':0, '106dB':0, '109dB':0, '112dB':0, '115dB':0, '118dB':0}
    runTime=0
    plot1.clear()
    x=[]
    y=[]
rec_start=None
def record_frame():
    global recording, rec_start, db, f
    time_trunc='%.1f' % (time.time()-rec_start)
    db_str='%.1f' % db
    f.write(time_trunc+','+db_str+'\n')
    if recording:
        win.after(500, record_frame)
    else:
        f.close()
def record():
    global recording, f, rec_start
    if not recording:
        rec.config(text='Stop')
        rec_start=time.time()
        recording=True
        f=open(os.getenv('HOMEPATH')+'\\Music\\headphone_levels.csv', 'w')
        f.write('Seconds,Decibels\n')
        record_frame()
    else:
        recording=False
        rec.config(text='Record')
ev = IAudioEndpointVolume.get_default()
win=Tk()
win.title('Headphone Levels')
win.grid()
win.iconbitmap(get_resource_path('snd.ico'))
win.resizable(False, False)
dosi_enabled_first=dosi_enabled
tabControl = Notebook(win)
root=Frame(tabControl)
sub=Frame(tabControl)
sub1=Frame(tabControl)
notify_timestamp=0
tabControl.add(root, text ='Meter')
measure=False
tabControl.add(sub, text ='Dosimeter')
tabControl.add(sub1, text ='Recording')
tabControl.pack(expand = 1, fill ="both")
gaugedb=SevenSegmentDigits(root, digits=3, digit_color='#00ff00', background='black')
gaugedb.grid(column=1, row=1)
graphdb=SevenSegmentDigits(sub1, digits=3, digit_color='#00ff00', background='black')
graphdb.grid(column=1, row=2)
Hovertip(gaugedb,'Current dB level')
root.limiter=85
gr_start=time.time()
led = Led(root, size=20)
led.grid(column=2, row=14)
Label(sub, text='Instantaneous dB level').grid(column=1, row=1)
Label(sub1, text='Instantaneous dB level').grid(column=1, row=1)
style=Style(win)
style.configure("1.Horizontal.TProgressbar", background='green')
style.configure("2.Horizontal.TProgressbar", background='yellow')
style.configure("3.Horizontal.TProgressbar", background='red')
style = Style(root)
style.layout('text.Horizontal.TProgressbar',
             [('Horizontal.Progressbar.trough',
               {'children': [('Horizontal.Progressbar.pbar',
                              {'side': 'left', 'sticky': 'ns'})],
                'sticky': 'nswe'}),
              ('Horizontal.Progressbar.label', {'sticky': ''})])
style.configure('text.Horizontal.TProgressbar', text='0 %')
db_levels=[82, 85, 88, 91, 94, 97, 100, 103, 106, 109, 112, 115, 118]
niosh_limits=[57600, 28800, 14400, 7200, 3600, 1800, 900, 450, 225, 120, 60, 30, 15]
db_levels.reverse()
runTime=0
niosh_limits.reverse()
enable_dosi=Checkbutton(sub, text='Enable Dosimeter (restart app to apply)', command=toggle_dosi)
enable_dosi.grid(column=1, row=3)
enable_dosi.state(['!alternate'])
if not reg_present:
    enable_dosi.config(state=DISABLED)
    enable_dosi.state(['!alternate'])
if dosi_enabled:
    enable_dosi.state(['selected'])
    Button(sub, text='Reset', command=reset).grid(column=2, row=3)
    Label(sub, text='Dose:', width=40).grid(column=1, row=4)
    timeLabel=Label(sub, text='0 sec', width=30)
    timeLabel.grid(column=2, row=2)
    dosebar=Progressbar(sub, maximum=100, mode='determinate', length=200, style='text.Horizontal.TProgressbar')
    dosebar.grid(column=2, row=4)
    pdose=Label(sub, text='Projected dose: 0 sec', width=40)
    for i in range(len(db_levels)):
        db_level=db_levels[i]
        exec("label_"+str(db_level)+"=Label(sub, width=40, text='"+str(db_level)+"dBA: 0/"+str(niosh_limits[i])+" sec')")
        exec("label_"+str(db_level)+".grid(column=1, row="+str(i+5)+")")
        exec("bar_"+str(db_level)+'=Progressbar(sub, length=200, style="1.Horizontal.TProgressbar", mode="determinate")')
        exec("bar_"+str(db_level)+'.grid(column=2, row='+str(i+5)+')')
    Label(sub1, text='Recording to Music\\headphone_levels.csv.\nMake sure the file does not exist, or it will be overwritten.').grid(column=1, row=3)
    rec=Button(sub1, text='Record', command=record)
    rec.grid(column=1, row=4)
else:
    Label(sub, text='Dosimeter is not enabled', width=60).grid(column=1, row=4)
    Label(sub1, text='Dosimeter must be enabled', width=60).grid(column=1, row=4)
Hovertip(led,'1 dB\nAll is OK')
led0 = Led(root, size=20)
led0.grid(column=2, row=13)
Hovertip(led0,'10 dB\nAll is OK')
led1 = Led(root, size=20)
led1.grid(column=2, row=12)
Hovertip(led1,'20 dB\nAll is OK')
led2 = Led(root, size=20)
led2.grid(column=2, row=11)
Hovertip(led2,'30 dB\nAll is OK')
led3 = Led(root, size=20)
led3.grid(column=2, row=10)
Hovertip(led3,'40 dB\nAll is OK')
led4 = Led(root, size=20)
led4.grid(column=2, row=9)
Hovertip(led4,'50 dB\nAll is OK')
led5 = Led(root, size=20)
led5.grid(column=2, row=8)
Hovertip(led5,'60 dB\nAll is OK')
led6 = Led(root, size=20)
led6.grid(column=2, row=7)
Hovertip(led6,'70 dB\nAll is OK')
led7 = Led(root, size=20)
led7.grid(column=2, row=6)
Hovertip(led7,'80 dB\nA little loud, may cause hearing damage in sensitive people.')
led8 = Led(root, size=20)
led8.grid(column=2, row=5)
Hovertip(led8,'90 dB\nLoud; repeated and/or long term exposure to this level may damage hearing.')
led9 = Led(root, size=20)
led9.grid(column=2, row=4)
Hovertip(led9,'100 dB\nCritically loud, even short exposure to this level can damage hearing.')
led10 = Led(root, size=20)
led10.grid(column=2, row=3)
Hovertip(led10,'110 dB\nDangerous, even short exposure to this level can damage hearing.')
led11 = Led(root, size=20)
led11.grid(column=2, row=2)
Hovertip(led11,"120 dB\nDangerous, even short exposure to this level can damage hearing.\nYou might feel pain at this level.")
Label(root, text='120').grid(column=1, row=2)
Label(root, text='100').grid(column=1, row=4)
Label(root, text='Danger').grid(column=3, row=4)
Label(root, text='80').grid(column=1, row=6)
Label(root, text='Loud').grid(column=3, row=6)
Label(root, text='60').grid(column=1, row=8)
Label(root, text='40').grid(column=1, row=10)
Label(root, text='20').grid(column=1, row=12)
Label(root, text='dB').grid(column=1, row=14)
Label(root, text='OK').grid(column=3, row=14)
Label(root, text='Max').grid(column=3, row=0)
Label(root, text='dB').grid(column=1, row=0)
Label(root, text='-').grid(column=1, row=3)
Label(root, text='-').grid(column=1, row=5)
Label(root, text='-').grid(column=1, row=7)
Label(root, text='-').grid(column=1, row=9)
Label(root, text='-').grid(column=1, row=11)
Label(root, text='-').grid(column=1, row=13)
Label(root, text='Volume Limit').grid(column=2, row=0)
dosimeter_times={'82dB':0, '85dB':0, '88dB':0, '91dB':0, '94dB':0, '97dB':0, '100dB':0, '103dB':0, '106dB':0, '109dB':0, '112dB':0, '115dB':0, '118dB':0}
sb=Spinbox(root, width=8, state='readonly', values=['75 dB', '80 dB', '85 dB', '90 dB', '95 dB', '100 dB', 'Off'], command=update_limiter)
sb.grid(column=2, row=1)
sb.set('85 dB')
dosidb=SevenSegmentDigits(sub, digits=3, digit_color='#00ff00', background='black')
dosidb.grid(column=1, row=2)
db=0
CHECK_FREQ=200
last_check=0
maxdb_display=SevenSegmentDigits(root, digits=3, digit_color='#00ff00', background='black')
maxdb_display.grid(column=3, row=1)
Hovertip(maxdb_display,'Max dB level while app was running')
appclosed=False
NUMERATOR, DENOMINATOR = A_weighting(RATE)
Hovertip(sb,'Maximum loudness before volume is reduced automatically')
max_decibel=0
os.path.expanduser(os.getenv('USERPROFILE'))+'\\Music\\myfile.txt'
x=[]
y=[]
def returnSum(myDict):
    mylist = []
    for i in myDict:
        mylist.append(myDict[i])
    final = sum(mylist)
    return final
def change_color(index):
    global db
    bar=db_levels[index]
    temp=dosimeter_times[str(bar)+'dB']/niosh_limits[index]
    if temp>=0.9:
        exec("bar_"+str(db_level)+".config(style='3.Horizontal.TProgressbar')")
    elif temp>=0.5:
        exec("bar_"+str(db_level)+".config(style='2.Horizontal.TProgressbar')")
    else:
        exec("bar_"+str(db_level)+".config(style='1.Horizontal.TProgressbar')")
def timer_dosi():
    global db, runTime, CHECK_FREQ, gr_start, start_check
    runTime+=time.time()-start_check
    x.append(time.time()-gr_start)
    y.append(db)
    if db>=82 and db<85:
        dosimeter_times['82dB']+=time.time()-start_check
    if db>=85 and db<88:
        dosimeter_times['85dB']+=time.time()-start_check
    if db>=88 and db<91:
        dosimeter_times['88dB']+=time.time()-start_check
    if db>=91 and db<94:
        dosimeter_times['91dB']+=time.time()-start_check
    if db>=94 and db<97:
        dosimeter_times['94dB']+=time.time()-start_check
    if db>=97 and db<100:
        dosimeter_times['97dB']+=time.time()-start_check
    if db>=100 and db<103:
        dosimeter_times['100dB']+=time.time()-start_check
    if db>=103 and db<106:
        dosimeter_times['103dB']+=time.time()-start_check
    if db>=106 and db<109:
        dosimeter_times['106dB']+=time.time()-start_check
    if db>=109 and db<112:
        dosimeter_times['109dB']+=time.time()-start_check
    if db>=112 and db<115:
        dosimeter_times['112dB']+=time.time()-start_check
    if db>=115 and db<118:
        dosimeter_times['115dB']+=time.time()-start_check
    if db>=118:
        dosimeter_times['118dB']+=time.time()-start_check
    start_check=time.time()
    win.after(CHECK_FREQ, timer_dosi)
def listen(old=0, error_count=0, min_decibel=100):
    global max_decibel
    global appclosed, plot1, db, runTime, notify_timestamp, last_check, CHECK_FREQ
    if not appclosed:
        try:
            try:
                block = stream.read(CHUNK)
            except (IOError, NameError) as e:
                if not appclosed:
                    error_count += 1
                    messagebox.showerror("Error, ", " (%d) Error recording: %s" % (error_count, e))
                    close()
            else:
                decoded_block = numpy.frombuffer(block, numpy.int16)
                y = lfilter(NUMERATOR, DENOMINATOR, decoded_block)
                if ev.GetMasterVolumeLevelScalar()>0:
                    new_decibel = (1-ev.GetMute())*(20*numpy.log10(rms_flat(y))+40+ev.GetMasterVolumeLevel())
                else:
                    new_decibel=0
                if new_decibel<0:
                    new_decibel=0
                if runTime>0:
                    runt=runTime
                else:
                    runt=0.1
                gaugedb.set_value(str(int(new_decibel)))
                dosidb.set_value(str(int(new_decibel)))
                graphdb.set_value(str(int(new_decibel)))
                if float(new_decibel)>float(max_decibel):
                    max_decibel=new_decibel
                maxdb_display.set_value(str(int(float(str(max_decibel)))))
                db=new_decibel
                if int(db)>0:
                    led.to_green(on=True)
                else:
                    led.to_grey(on=True)
                for i in range(0, 12):
                    if int(new_decibel)>=(10*(i+1)):
                        if i>=9:
                            exec("led"+str(i)+".to_red(on=True)")
                        elif i>=7:
                            exec("led"+str(i)+".to_yellow(on=True)")
                        else:
                            exec("led"+str(i)+".to_green(on=True)")
                    else:
                        exec("led"+str(i)+".to_grey(on=True)")
                if new_decibel>root.limiter:
                    keyboard.tap(Key.media_volume_down)
                if new_decibel>=100 and time.time()-notify_timestamp>=180:
                    notify_timestamp=time.time()
                    toaster.show_toast('Turn that shit down!', 'Headphone levels have exceeded 100 dB. Listening at this level for even a short time can cause permanent hearing damage.', duration=5, icon_path=None, threaded=True)
                if dosi_enabled_first:
                    percent=(returnSum(dosimeter_times)/runt)*100
                    dosebar['value']=float(percent)
                    style.configure('text.Horizontal.TProgressbar',text='%.1f %%' % percent)
                    pdose.config(text='Projected dose: '+str(int((returnSum(dosimeter_times)*8)/runt))+' sec')
                    timeLabel.config(text=str(int(runt))+' sec')
                    for i in range(len(db_levels)):
                        db_level=db_levels[i]
                        exec("label_"+str(db_level)+".config(text='"+str(db_level)+"dB: "+str(int(dosimeter_times[str(db_level)+'dB']))+'/'+str(niosh_limits[i])+" sec')")
                        exec("bar_"+str(db_level)+"['value']="+str((dosimeter_times[str(db_level)+'dB']/niosh_limits[i])*100))
                        change_color(i)
            win.after(50, listen)
        except TclError:
            pass
win.protocol('WM_DELETE_WINDOW', close)
start_check=time.time()
if __name__ == '__main__':
    if not appclosed:
        messagebox.showwarning('Warning', 'This meter was calibrated for the HP Laptop 15t-dy100 and EarPods. As every headphone/earbud and device has a different power/decibel output, other headphones and/or computers may not be accurate. Accuracy depends on ear placement. Also, split your output between your headphones and the VB Cable.')
    start=time.time()
    if dosi_enabled:
        timer_dosi()
    listen()
    if appclosed:
        win.destroy()
    try:
        win.mainloop()
    except TclError:
        pass
