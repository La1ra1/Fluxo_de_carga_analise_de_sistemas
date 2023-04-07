import win32com.client
import numpy as np
import matplotlib.pyplot as plt
import openpyxl as op
from os import getcwd

class DSS:

    def __init__(self, local):

        self.local = local

        self.dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")

        if self.dssObj.Start(0) is False:
            print("Dss Failed to Start")
        else:
            self.dssText = self.dssObj.Text
            self.dssCircuit = self.dssObj.ActiveCircuit
            self.dssSolution = self.dssCircuit.solution
            self.dssCktElement = self.dssCircuit.ActiveCktElement
            self.dssText.Command = 'Compile [' + str(local) + ']'
            self.dssText.Command = 'solve'
            self.loads = self.dssCircuit.Loads

bus13 = DSS(getcwd() + "\IEEE13Nodeckt.dss")

bus13.dssText.Command = 'Redirect Loadshapes.dss'

wb = op.load_workbook('Cargas.xlsx')

tupla=wb['Plan1']['B3:C17']

ls_info = {}

for data in tupla:
    ls_info.update({str(data[0].value) : data[1].value})

bus13.loads.First
for i in range(bus13.loads.count):
    bus13.loads.daily = ls_info[bus13.loads.Name]
    bus13.loads.Next

bus13.loads.First
for i in range(bus13.loads.count):
    print(bus13.loads.Name)
    bus13.loads.Next

bus13.dssText.Command = 'Set mode=daily'
bus13.dssText.Command = 'Set stepsize=15m'
bus13.dssText.Command = 'Set number=1'

losses = []
total_power = []

for i in range(96):
    bus13.dssSolution.Solve()
    total_power.append(bus13.dssCircuit.TotalPower[0]/1000*-1)
    losses.append(bus13.dssCircuit.Losses[0]/1000)

bus13.loads.First
for i in range(bus13.loads.count):
    bus13.loads.model = 1
    bus13.loads.Next

losses_r = []
total_power_r = []

for i in range(96):
    bus13.dssSolution.Solve()
    total_power_r.append(bus13.dssCircuit.TotalPower[0]/1000*-1)
    losses_r.append(bus13.dssCircuit.Losses[0]/1000)

bus13.loads.First
for i in range(bus13.loads.count):
    bus13.loads.model = 2
    bus13.loads.Next

losses_s = []
total_power_s = []

for i in range(96):
    bus13.dssSolution.Solve()
    total_power_s.append(bus13.dssCircuit.TotalPower[0]/1000*-1)
    losses_s.append(bus13.dssCircuit.Losses[0]/1000)

bus13.loads.First
for i in range(bus13.loads.count):
    bus13.loads.model = 5
    bus13.loads.Next

losses_t = []
total_power_t = []

for i in range(96):
    bus13.dssSolution.Solve()
    total_power_t.append(bus13.dssCircuit.TotalPower[0]/1000*-1)
    losses_t.append(bus13.dssCircuit.Losses[0]/1000)

s = 0
print('Demanda max: ' + str(max(total_power_s)))
print('Demanda média: ' + str(np.mean(total_power_s)))

print('FC: ' + str(np.mean(total_power_s)/max(total_power_s)))

#plt.plot(np.arange(0, 24, 0.25),losses,color="#1f77b4")
plt.plot(np.arange(0, 24, 0.25),losses_r,color='Green')
plt.plot(np.arange(0, 24, 0.25),losses_s,color='Red')
plt.plot(np.arange(0, 24, 0.25),losses_t,color='Black')

plt.xlabel("Hora")
plt.ylabel("Perdas Ativas (W)")
plt.title("Curva de Perda")
plt.xlim([0, 24])
plt.xticks(range(0, 25, 4))
plt.legend(['Cargas de Potência Constante','Cargas de Impedância Constante', 'Cargas de Corrente constante'])

plt.show()
