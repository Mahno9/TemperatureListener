import os

import serial
import serial.tools.list_ports
import sys
import time
import datetime
import csv


def SplitLine(line: str) -> list:
    expectedPrefix = "=9600A71"
    if not line.startswith(expectedPrefix):
        print("Warning: unexpected line prefix in next line: {}".format(line))
        return [line, -1, "???"]
    prefLen = len(expectedPrefix)
    return [line[:prefLen], int(line[prefLen:prefLen + 4]), line[prefLen + 4:]]


def ProcessSerialDataLine(line):
    dataRow = SplitLine(line)
    curdt = datetime.datetime.fromtimestamp(time.time()).isoformat()
    print(curdt, dataRow)
    return [curdt] + dataRow


def ListendCOMPort(port: str):
    data = []
    needStopThread = False

    def ListendFunc():
        serialPort = serial.Serial(port=port, baudrate=9600,
                                   bytesize=8, timeout=2, stopbits=1)
        while not needStopThread:
            if (serialPort.in_waiting > 0):
                serialString = serialPort.read_all()
                for line in serialString.decode('utf8', errors='ignore').split('\n'):
                    if len(line) == 0:
                        continue
                    data.append(ProcessSerialDataLine(line))
        else:
            serialPort.close()

    import threading
    thread = threading.Thread(target=ListendFunc())
    thread.start()
    sys.stdin.read(1)
    needStopThread = True
    thread.join()

    return ["Time", "Prefix", "Temperature", "Postfix"], data


def WriteToCSV(dataRowsNames, dataRows, filename):
    with open(filename, 'w', newline='') as outFile:
        writer = csv.writer(outFile, dialect='excel')
        writer.writerow(dataRowsNames)
        writer.writerows(dataRows)


def main():
    ports = serial.tools.list_ports.comports()

    idx = 0
    for port, desc, hwid in sorted(ports):
        idx += 1
        print("{}. {}: {} [{}]".format(idx, port, desc, hwid))

    targetPortIdx = int(input("Enter the number of COM port info line: ")) - 1
    dataRowsNames, dataRows = ListendCOMPort(ports[targetPortIdx][0])

    tempOutFilename = "Result.csv"
    WriteToCSV(dataRowsNames, dataRows, tempOutFilename)
    while 1:
        resultFilename = input("Enter a name for the result .csv file (without extension): ") + ".csv"
        if os.path.exists(resultFilename):
            print("File with this name already exists: {}. Please, enter another one.".format(resultFilename))
            continue
        break
    os.rename(tempOutFilename, resultFilename)
    print("Result filename is: {}".format(resultFilename))


if __name__ == '__main__':
    main()
