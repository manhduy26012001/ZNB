import pyvisa

adress = "RSNRP::0x0016::103177::INSTR"
rm = pyvisa.ResourceManager()
instrusmen = rm.open_resource(adress)
instrusmen.timeout(30000)
idn1 = instrusmen.query('*IDN?')
print(idn1)