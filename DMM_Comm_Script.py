import pyvisa
import time

rm = pyvisa.ResourceManager()
instr = rm.open_resource("TCPIP0::192.168.0.11::45454::SOCKET")
instr.timeout = 10000

instr.read_termination = '\r\n'
instr.write_termination = '\n'

try:
    instr.write("*CLS")
    instr.write("*RST")
    instr.write("SYST:REM")
    instr.write("FUNC 'VOLT:DC'")
    time.sleep(0.5)

    # THE TEST:
    result = instr.query("VAL1?")
    print("VAL1?:", result.strip())

    result2 = instr.query("SENS:DATA?")
    print("SENS:DATA?:", result2.strip())

except Exception as e:
    print("Error:", e)

finally:
    instr.close()
