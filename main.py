# # TODO - here we go I suppose :(

from sys import byteorder
from tracemalloc import start
import gspread
import serial
import time
import datetime
import colorama as clr
from colorama import Fore, Style


# seems to be a shutdown code to the device

# we need to send an initialization message!

# fa 06 00 00 00 00 00 00 06 f8

init_message = 0xfa0500000000000005f8 
close_message = 0xfa0600000000000006f8

# INSTALL COMMAND: pip install gspread pyserial colorama

# google sheet: https://docs.google.com/spreadsheets/d/1mdjazOoiNfkSh2w03e4G9gWg25xuktNSD1W5mzzhZc4/edit?usp=sharing

# recommend having this sheet open! :) ^^


# 791 15/08/2025 14:43:14 IRP_MJ_WRITE UP STATUS_SUCCESS fa 05 00 00 00 00 00 00 05 f8  ú........ø 10  COM6  
# 842 15/08/2025 14:44:16 IRP_MJ_WRITE UP STATUS_SUCCESS fa 05 00 00 00 00 00 00 05 f8  ú........ø 10  COM6  


# Settings for running program ---------------------------------------------------

run_flag=True 
increment_loc=0
clr.init(autoreset=True)

# --------------------------------------------------------------------------------

# Settings for google sheet ------------------------------------------------------

sa = gspread.service_account(filename="C:\\Users\\SEVT-1\\Desktop\\impedance-testing-cells-2025-97102ef1723c.json")
sh = sa.open("Battery Impedances INR21700-53G")

wks = sh.worksheet("raw_impedance")

def find_next_id():
    """returns the next id for the cell test"""
    values_list = wks.col_values(1)
    print(values_list)
    # only returns filled values
    return len(values_list)
# --------------------------------------------------------------------------------

# SETTINGS FOR SERIAL  -----------------------------------------------------------
PORT = "COM5"
BAUDRATE=9600
ser = serial.Serial(PORT, BAUDRATE, bytesize=serial.EIGHTBITS,
                    parity=serial.PARITY_EVEN,
                    stopbits=serial.STOPBITS_ONE,
                    timeout=2,
                    xonxoff=True,
                    rtscts=False,
                    dsrdtr=False)
ser.dtr=True
ser.rts=True
ser.reset_input_buffer()
ser.reset_output_buffer()
time.sleep(0.1)
ser.flush() 



TEST_CURRENT = 5 # Amps
test_write_messages = [0xfa090064000000006df8,
                       0xfa0900c800000000c1f8,
                       0xfa09013c0000000034f8,
                       0xfa0901a000000000a8f8,
                       0xfa090214000000001ff8]

format_init_message = init_message.to_bytes(10, byteorder='big')
ser.write(format_init_message)
time.sleep(1)
if ser.in_waiting:
    ser.read(ser.in_waiting)

print("was it printed?")

debounce_time = 3 # seconds
num_tests = 5
# start_byte = 0xfa
# end_byte = 0xf8


def get_voltage(data):
    """ extract voltage from data """
    processed_data = int.from_bytes(data, byteorder='big')
    voltage_msg=((processed_data>>104) & 0xFFFF) / 1000.000
    return voltage_msg



def get_resistance(data, current):
    """ extract resistance from data """
    processed_data = int.from_bytes(data, byteorder='big')
    current_msg=(processed_data>>88) & 0xFFFF
    current_resistance = float(current_msg) / float(current)
    return current_resistance


def collect_data(test_current):
    """returns a list of 5 data points, all """
    today = datetime.date.today()
    voltage_list = [0.000]*num_tests
    data_list = [0.0]*7 

    # loc 5 holds the average, loc 6 holds avg test voltage, loc 7 holds date
    test_message = test_write_messages[test_current-1]
    format_test_message = test_message.to_bytes(10, byteorder='big')
    for test_number in range(num_tests):
        while True:
            if ser.in_waiting:
                voltage_reading=ser.read(ser.in_waiting)
                voltage_list[test_number] = get_voltage(voltage_reading)
                break


        ser.reset_input_buffer()
        written_data=ser.write(format_test_message)
        ser.reset_input_buffer()

        while True:
            # let's wait for a message
            if ser.in_waiting:
                data = ser.read(ser.in_waiting)
                print(data)
                break
            time.sleep(0.001)
        
        data_list[test_number] = get_resistance(data, test_current)
        time.sleep(debounce_time) # adjust for how long it takes for voltage to recover

 
    for index, value in enumerate(voltage_list):
        data_list[6]+=value/float(num_tests)# gets average voltage
        data_list[5] += data_list[index] / float(num_tests) #gets average resistance

    data_list+=[today.strftime('%m/%d/%Y')]
    # data has now been collected, 
    # send to review and see if should be written?
    return data_list # returns data with 5 points, average in 6, voltage average in 7, and date in 8

# --------------------------------------------------------------------------------


starting_ID = find_next_id() 
print(starting_ID)


# wks.update([[1, 2], [3, 4]], 'A1:B2')
try:
    while run_flag:
        print(Fore.GREEN + Style.BRIGHT + f"TESTING CELL #{starting_ID+increment_loc}")
        user_input = input(Fore.YELLOW + Style.BRIGHT + "Press ENTER to test, or Q to quit\n")

        if user_input=="":
            print(Fore.GREEN + Style.BRIGHT + "TESTING...")
        elif user_input.strip().lower()=="q":
            print(Fore.RED + Style.BRIGHT + "Quitting...")
            ser.close()
            run_flag=False
            break

        collected_data = collect_data(TEST_CURRENT)

        for index in range(num_tests):
            print(f"Resistance {index}: {collected_data[index]}\n")
        
        user_input = input(Fore.YELLOW + Style.BRIGHT + "Press ENTER to verify data, or 'R' to retest\n")
        if user_input.strip().lower() == 'r':
            continue # should hopefully restart the while loop
        # print("Saving...\n")
        pos = starting_ID + increment_loc
        # added the round feature to make writing easier
        print(Fore.RED + Style.BRIGHT + f"CELL # {pos} : {round(collected_data[5],1)}mR \n\n")
        custom_loc = f'A{pos}:I{pos}'
        wks.update([[pos]+collected_data], custom_loc)
        increment_loc+=1
except KeyboardInterrupt:
    format_close_message = close_message.to_bytes(10, byteorder='big')
    ser.write(format_close_message)
    ser.close()
    







# **********************************************************************

    # the code should continue running here
        # tell the user what number to write on the cell 

    # get the data 

    # validate the data 

    # write the data - request to press enter 
        # tell user what value to write on the cell

    # wait for next cell  - request to press enter again to begin collecting 
# custom_loc = f'A{starting_ID}:D{starting_ID}'

# print("Rows:", wks.row_count)
# print("Cols: ", wks.col_count)

# print(wks.acell("A9").value)
# print(wks.cell(3,4).value) # down 3 rows, over 4 col



# there is a .update function 
# which works for a group of cells!, will probably use this 
# to update 5 cells at once

# we want this code to check what the latest added value was 
# (only check left-most column for this)
# we can do this by iterating through each row in row_count,
    # and if none of them are empty, we create a new row, putting
    # the next number there.

# performance likely the best if I grab all the data from column 1
# and just see how large it currently is, then use that to what value 
# we should be continuing from the latest, and then in formatted text, ask if it looks good


# data = collect_data(TEST_CURRENT)



























# # let's get this shawty started
# import pyvisa 
# import time

# # setup our user params
# I0 = 0.0 # default current
# I1 = 0.5 # low current
# I2 = 1.0 # high current
# dwell_time = 0.6 # 300ms
# rest_time = 1.0 # 1 second
# settling_time = 0.02 # let things stabilize before reading
# repeat_count = 5 # repeat it 5 times


# # connecting to RIGOL !!! _-----------------------
# rm = pyvisa.ResourceManager()
# load = rm.open_resource('USB0::0x1AB1::0x0E11::DL3D254300331::INSTR',) 
# # ^^ put our actual serial resource in the dashes for the USB-TMC connection

# # could also print available resources 
# # print(rm.list_resources())

# def measure_voltage():
#     time.sleep(settling_time)
#     v = float(load.query(":MEAS:VOLT?"))
#     return v

# try: 
#     # begin this
#     load.write("*RST")
#     load.write(":SENS:REM ON")
#     load.write(":MODE CURR")
#     load.write(":INPUT ON")

#     time.sleep(2)
#     print(f"MEASURED voltage {float(load.query(":MEAS:VOLT?"))}")

#     impedances = []

#     for i in range(repeat_count):
#         load.write(f":CURR {I1}")
#         time.sleep(dwell_time)
#         v1 = measure_voltage()
#         c1 = float(load.query(":MEAS:CURR?"))
#         print(f"CURR I1: {c1}")

#         # set I2
#         load.write(f":CURR {I2}")
#         time.sleep(dwell_time)
#         v2 = measure_voltage()
#         c2 = float(load.query(":MEAS:CURR?"))
#         print(f"CURR I2: {c2}")


#         delta_v = v2-v1
#         delta_i = c2-c1

#         z = delta_v/delta_i
#         print(f"[Cycle {i+1}] deltaV = {delta_v:.4f} V, Z = {z*1000:.2f} mOhms")
#         impedances.append(z)
#         # load.write(f":CURR {I0}")
#         # time.sleep(rest_time)

#     avg_z = sum(impedances)/len(impedances)
#     print(f"\nEstimated Internal Resistance: {avg_z*1000:.2f} mOhms")
# finally:
#     load.write(":INPUT OFF")
#     load.close()

