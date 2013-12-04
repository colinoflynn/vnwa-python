import win32api, win32ui, win32gui, win32con
import subprocess
from multiprocessing import Process
from subprocess import Popen
DETACHED_PROCESS = 0x00000008
import sys
import time

class VNWAConnector(object):
    def __init__(self):
        message_map = {
            win32con.WM_USER: self.wndProc
        }
        wc = win32gui.WNDCLASS()
        wc.lpfnWndProc = message_map 
        wc.lpszClassName = 'VNWAListenerClass'
        hinst = wc.hInstance = win32api.GetModuleHandle(None)
        classAtom = win32gui.RegisterClass(wc)
        self.hwnd = win32gui.CreateWindow (
            classAtom,
            "VNWA Crappy Listener",
            0,
            0, 
            0,
            win32con.CW_USEDEFAULT, 
            win32con.CW_USEDEFAULT,
            0, 
            0,
            hinst, 
            None
        )
        #print self.hwnd

        self.waiting = True

    def wndProc(self, hwnd, msg, wparam, lparam):
        #print "%d: %d %d %d"%(hwnd, msg, wparam, lparam)

        self.waiting = False
        
        if (wparam & 0xffff0000) == 0:  
            self.VNWA_HWND = lparam    
            self.VNWA_MSG = wparam
            print "Connected to VNWA Process (%d, %d)"%(self.VNWA_MSG, self.VNWA_HWND)
        else:
            ecode = wparam >> 16

            if ecode == 1:
                #print "Command OK"
                pass
            elif ecode == 2:
                raise IOError("Script file error, or non-existent file")
            elif ecode == 3:
                raise IOError("File access error?")
            elif ecode == 5:
                raise IOError("VNWA Overload - check audio settings")
            else:
                raise ValueError("Unknown error = %d"%ecode)

    def sendMessage(self, wparam, iparam, wait=True):
        self.waiting = True
        win32api.PostMessage(self.VNWA_HWND, self.VNWA_MSG, wparam, iparam)

        if wait:
            self.waitResponse()

    def waitResponse(self, timeout=20):

        for i in range(0,timeout*100):    
            win32gui.PumpWaitingMessages()
            if self.waiting == False:
                break
            time.sleep(0.01)

    def setRFile(self, rstring):
        self.sendMessage(6, 0)
        for r in rstring:
            self.sendMessage(6, ord(r))

    def setWFile(self, rstring):
        self.sendMessage(7, 0)
        for r in rstring:
            self.sendMessage(7, ord(r))
            

class VNWA(object):

    def startVNWA(self, exeloc):
        self.vnaconn = VNWAConnector()

        cmd = [
            exeloc,
            '-remote',
            '-callback',
            str(self.vnaconn.hwnd),
            str(win32con.WM_USER),
            '-debug'
          ]

        processVNWA = Popen(cmd,shell=False,stdin=None,stdout=None,stderr=None,close_fds=True,creationflags=DETACHED_PROCESS)

        self.vnaconn.waitResponse()

    def closeVNWA(self):
        """ Terminate the VNWA Program """
        self.vnaconn.sendMessage(0, 0)

    def sweep(self, S21=False, S11=False, S12=False, S22=False):
        """ Do a Sweep """
        swpmode = 0
        if S21:
            swpmode |= 1<<0

        if S11:
            swpmode |= 1<<1

        if S12:
            swpmode |= 1<<2

        if S22:
            swpmode |= 1<<3
        
        self.vnaconn.sendMessage(1, swpmode)

    def loadCal(self, filename):
        self.vnaconn.setRFile(filename)
        self.vnaconn.sendMessage(2, 0)

    def loadMasterCal(self, filename):
        self.vnaconn.setRFile(filename)
        self.vnaconn.sendMessage(3, 0)

    def writeS2P(self, filename):
        self.vnaconn.setWFile(filename)
        self.vnaconn.sendMessage(4, 0)

    def setStartFreq(self, freq):
        self.vnaconn.sendMessage(8, freq)

    def setStopFreq(self, freq):
        self.vnaconn.sendMessage(9, freq)

    def setTXPower(self, power):
        """ Set TX Power in range 0...16383 """
        self.vnaconn.sendMessage(17, power)
    

def main():
    vna = VNWA()

    vna.startVNWA('C://E//Documents//VNA Stuff//vnwa//VNWA.exe')

    vna.setStartFreq(10E6)
    vna.setStopFreq(100E6)
    vna.loadCal('C:\ctest.cal')
    vna.sweep(S21=True)

if __name__ == "__main__":
    main()
