import serial
import time
from datetime import datetime
from pathlib import Path
import os
import sys
import threading
import concurrent.futures as futures
import subprocess

FILENAME = 'dustdata.csv'

def current_time():
    """get current time in nice format as str"""
    return datetime.now().strftime("%Y-%m-%d %H:%M")

def _clean_b(b):
    """use to turn binary to float"""
    s = b.decode().rstrip()
    f = float(s)
    return f

def _read_ser(ser):
    """attemp to readline from open serial port, ser is open ser object"""
    b = ser.readline()
    f = _clean_b(b)
    return f

def _PS_read(COM):
    print('Executing directly in Powershell subprocess')
    psloc = r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
    cmd = f'{psloc} $port= new-Object System.IO.Ports.SerialPort COM{COM},9600,None,8,one; $port.open(); $port.readline(); $port.close()'
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, shell=True)
    b = proc.stdout.read()
    proc.stdout.close()
    f = _clean_b(b)
    return f

def _timout(func):
    """executes func with a max timeout"""
    with futures.ThreadPoolExecutor(max_workers=1) as executor:
        future = executor.submit(func)
        try:
            resp = future.result(30)
        except futures.TimeoutError:
            print('Connection timed out')
            resp = None
        else:
            print(resp)
        executor._threads.clear()
        futures.thread._threads_queues.clear()
        return resp

def collect_now(COM):
    """read every 2 seconds and return value as float"""
    ser = serial.Serial(f'COM{COM}',9600)
    time.sleep(2)
    try:
       f = _timout(_read_ser(ser))
    except:
        print('Could not connect')
        f = None
    finally:
        ser.close()
    return f

def save_val(filename, val):
    """saves date and val to file in current folder"""
    fp = Path(__file__).resolve().parent.joinpath(filename)
    #print(f'before checking if exists:   {fp}')
    file_exists = os.path.isfile(str(fp)) 
    if file_exists:
        with open(str(fp), 'a') as f:
            f.write(val)
    else:
        fp.touch()

if __name__ == '__main__':
    COM = input('Enter COM port integer: ')
    while True:
        if collect_now(COM):
            f = collect_now(COM)
        else:
            f = _PS_read(COM)
        val = f'{current_time()}, {str(f)}\n'
        print(val)
        #val = '0.63, 2000-01-02 12:33\n'
        save_val(FILENAME, val)