from __future__ import print_function
from gsmmodem.modem import GsmModem, SentSms
import xlwings as xw
import  os

import serial
import time



import logging



# PORT = 'COM5' # ON WINDOWS, Port is from COM1 to COM9 ,
# We can check using the 'mode' command in cmd

BAUDRATE = 115200

PIN = None  # SIM card PIN (if any)



def modifyMapping(mapping,tex:str,number:str):

    if "{name}" in tex:
        for val in mapping:
            if number in str(val[0]):
                tex=tex.replace("{name}",val[1])
                if len(tex)>150:
                    tex= tex[:150]
                return tex

    else:
        return tex







def send_sms(recipients, message,port,DNC,mappingarray):
    try:

        comport = port  # Change this to the correct COM port
        s = serial.Serial()
        s.port = comport
        s.baudrate = 9600
        # s.parity = serial.Parity.NONE
        # s.bytesize = serial.EIGHTBITS
        # s.stopbits = serial.STOPBITS.ONE
        # s.dtr = True
        # s.rts = True



        for index, recipient in enumerate(recipients):
            phone_number = str(recipient).replace("+", "0").replace(".0", "").strip()
            if not isINDNC(DNC,phone_number):
                try:
                    message=modifyMapping(mappingarray, message, phone_number)
                    s.open()
                    s.write(b"AT\r\n")
                    time.sleep(1)
                    s.write(b"AT+CMGF=1\r\n")
                    time.sleep(1)
                    s.write(b"AT+CMGS=\"" + phone_number.encode() + b"\"\r\n")
                    time.sleep(1)
                    s.write(message.encode() + b"\x1A\r\n")
                    time.sleep(1)
                    s.close()
                    print(f"Message Sended to {phone_number}")
                except Exception as e:
                    print(f"Error: {e}")


    except Exception as e:
        print("big error "+str(e))


def isINDNC(DNC,number):
    for val in DNC:
        # print(val)
        if number in str(val):
            return True

    return False
# Replace 'recipient_numbers' with a list of actual phone numbers and 'Your message here' with your message
if __name__=="__main__":
    print("Opening Excel")
    wb_path = os.path.join(os.getcwd(), "Recepiant.xlsm")
    wb = xw.Book(wb_path)
    # Access the sheets
    sheet = wb.sheets["Dash"]
    mapping = wb.sheets["Mapping"]
    DNC = wb.sheets["DNC"]

    print("Getting Data")
    last_cell = mapping.cells(1, 1).end("down").row
    mappingarray=mapping.range(f"A2:B{last_cell}"  ).value


    last_cell = DNC.cells(1, 1).end("down").row
    DNCNumbers=DNC.range(f"A2:A{last_cell}" ).value


    last_cell = sheet.cells(1, 1).end("down").row
    first_column_values = sheet.range(f"A2:A{last_cell}").value

    # Read data from the sheet

    print("Sending Sms")
    port = sheet.range("B2").value.strip()
    message = sheet.range("C2").value

    # print("present"  ,isINDNC(DNCNumbers,"123450998"))
    send_sms(first_column_values,message,port,DNCNumbers,mappingarray)
    s=input("Done !")

