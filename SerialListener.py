import serial
import serial.tools.list_ports
import sys
import time
import datetime

ports = serial.tools.list_ports.comports()

idx = 0
for port, desc, hwid in sorted(ports):
    idx += 1
    print("{}. {}: {} [{}]".format(idx, port, desc, hwid))

targetPortIdx = int(input("Enter the number of COM port info line: ")) - 1


def GetTemperature(rawValue: str):
    return int(rawValue.split("=9600A71")[1][:-3])


def ListendCOMPort(port:str):
    data = []
    needStopThread = False

    def ListendFunc():
        serialPort = serial.Serial(port=port, baudrate=9600,
                                   bytesize=8, timeout=2, stopbits=1)
        while not needStopThread:
            if (serialPort.in_waiting > 0):
                serialString = serialPort.read_all()
                temperature = GetTemperature(serialString.decode('utf8', errors='ignore'))
                curdt = datetime.datetime.fromtimestamp(time.time())
                print(curdt.isoformat(), temperature)
        else:
            serialPort.close()

    import threading
    thread = threading.Thread(target=ListendFunc())
    thread.start()
    sys.stdin.read(1)
    needStopThread = True
    thread.join()

    return data

from pprint import pprint
pprint(ListendCOMPort(ports[targetPortIdx][0]))