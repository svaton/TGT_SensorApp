import sys
import serial.tools.list_ports
port_num = []

class COM_setup:
    def serial_ports(self):
        ports = serial.tools.list_ports.comports()
        result=[]
        for port in sorted(ports):
            result.append(port.device + " "+ port.description)
            port_num.append(port.device)
##            print(port)
        return result, port_num

    def get_port(self,i):
       return port_num[i]

