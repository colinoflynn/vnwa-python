# VNWA Control from a Python Appliction
# Copyright 2013 Colin O'Flynn
#
# Released under MIT License:
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.


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

    def startVNWA(self, exeloc, debug=True):
        self.vnaconn = VNWAConnector()

        cmd = [
            exeloc,
            '-remote',
            '-callback',
            str(self.vnaconn.hwnd),
            str(win32con.WM_USER)            
          ]

        if debug:
            cmd.append('-debug')

        processVNWA = Popen(cmd,shell=False,stdin=None,stdout=None,stderr=None,close_fds=True,creationflags=DETACHED_PROCESS)

        self.vnaconn.waitResponse()

    def closeVNWA(self):
        """ Terminate the VNWA Program """
        self.vnaconn.sendMessage(0, 0)

    def sweepOnce(self, S21=False, S11=False, S12=False, S22=False):
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

    def sweepContinous(self, S21=False, S11=False, S12=False, S22=False):
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
        
        self.vnaconn.sendMessage(18, swpmode)

    def stopSweep(self, stopNow=False):
        """ Stop sweep """
        if stopNow:
            par = 1
        else:
            par = 0

        self.vnaconn.sendMessage(19, par)

    def loadCal(self, filename):
        """ Load cal file """
        self.vnaconn.setRFile(filename)
        self.vnaconn.sendMessage(2, 0)

    def loadMasterCal(self, filename):
        """ Load master cal file """
        self.vnaconn.setRFile(filename)
        self.vnaconn.sendMessage(3, 0)

    def writeS2P(self, filename):
        """ Write data to S2P File """
        self.vnaconn.setWFile(filename)
        self.vnaconn.sendMessage(4, 0)

    def setStartFreq(self, freq):
        """ Set sweet start frequency in Hz """
        self.vnaconn.sendMessage(8, freq)

    def setStopFreq(self, freq):
        """ Set sweep stop frequency in Hz """
        self.vnaconn.sendMessage(9, freq)

    def setTXPowerLinear(self, power):
        """ Set TX Power in range 0...16383 """
        self.vnaconn.sendMessage(17, power)

    def setTXPowerdBm(self, power):
        """ Set TX Power in range -67 ... -17 dBm """

        #Convert to linear
        pw = 10 ** (power / 20.0)

        #Scale by VNWA Constant
        pw = pw * 115981.4

        #Make integer
        pw = round(pw)
        if pw < 0:
            pw = 0
        if pw > 16383:
            pw = 16383

        print pw

        self.setTXPowerLinear(pw)
        

    def setRFFreq(self, freq):
        """ Set RF DDS Frequency in Hz """
        self.vnaconn.sendMessage(14, freq)

    def setLOFreq(self, freq):
        """ Set LO DDS Frequency in Hz """
        self.vnaconn.sendMessage(15, freq)

    def setVNWAFreq(self, freq):
        """ Set RF & LO DDS Frequency in Hz with IF offset """
        self.vnaconn.sendMessage(16, freq)
        
def main():
    vna = VNWA()

    vna.startVNWA('C://E//Documents//VNA Stuff//vnwa//VNWA.exe')

    vna.setStartFreq(10E6)
    vna.setStopFreq(100E6)
    vna.loadCal('C:\ctest.cal')
    vna.sweepOnce(S21=True)

    vna.setVNWAFreq(1E6)

    vna.setTXPowerdBm(-33)

if __name__ == "__main__":
    main()
