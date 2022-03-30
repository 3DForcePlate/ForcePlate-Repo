import serial
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import time
from time import perf_counter_ns as nano
import scipy.signal as signal
from scipy.signal import savgol_filter
from scipy.signal import butter
import xlsxwriter as xlsx

# creating workbook and sheets and adding athlete name
workbook = xlsx.Workbook('Force_Plate_Data.xlsx')
worksheet = workbook.add_worksheet()
name = input("Enter user's name how you want it to appear in excel: ")

# Would need to change port to the port your computer says USB is in    
ser = serial.Serial('/dev/tty.usbmodem14301', baudrate = 2000000, timeout =1)
res = 5/(pow(2,12)-1)

#Asking for user weight
pounds = float(input("Enter user weight in lbs: "))
kilo = pounds * 0.45359237
print("Weight entered in kg = ", kilo)
gravity = -9.81

#Create array for offsets and real data
dataOffset = np.array([])
dataOffsetH0 = np.array([])
dataOffsetH1 = np.array([])
dataOffsetH2 = np.array([])
dataOffsetH3 = np.array([])
dataOffsetV0 = np.array([])
dataOffsetV1 = np.array([])
dataOffsetV2 = np.array([])
dataOffsetV3 = np.array([])

data = np.array([])
dataH0 = np.array([])
dataH1 = np.array([])
dataH2 = np.array([])
dataH3 = np.array([])
dataV0 = np.array([])
dataV1 = np.array([])
dataV2 = np.array([])
dataV3 = np.array([])


iOffset = 0
i = 0

#Calculate offset that will be used for removal
def offset(nums):
    sum_num = 0
    for t in nums:
        sum_num = sum_num + t
    return sum_num /len(nums)

#Start time and tare
endOffsetCal = 800
end = 1000
start = nano()
print('Starting tare...')

#Start calculating offset after certain amount of samples
while iOffset < endOffsetCal:
    #until number of samples desired reached:
    #remove space that splits data
    featherData = ser.readline().decode('ascii')
    remove = list(map(float, featherData.split(' ')))
    for x in remove:
        withres = x*res
        dataOffset = np.append(dataOffset,withres)
        
    #Split data array into 8 seperate arrays based off what sensor the data comes from
    dataOffsetH0 = dataOffset[0::8]
    dataOffsetH1 = dataOffset[1::8]
    dataOffsetH2 = dataOffset[2::8]
    dataOffsetH3 = dataOffset[3::8]
    dataOffsetV0 = dataOffset[4::8]
    dataOffsetV1 = dataOffset[5::8]
    dataOffsetV2 = dataOffset[6::8]
    dataOffsetV3 = dataOffset[7::8]
    
    iOffset = iOffset + 1

#Find offset for each sensor    
avgH0 = offset(dataOffsetH0)
avgH1 = offset(dataOffsetH1)
avgH2 = offset(dataOffsetH2)
avgH3 = offset(dataOffsetH3)
avgV0 = offset(dataOffsetV0)
avgV1 = offset(dataOffsetV1)
avgV2 = offset(dataOffsetV2)
avgV3 = offset(dataOffsetV3)


print('Ready to Use')

while i < end:
    #until number of samples desired reached:
    #remove space that splits data
    featherData = ser.readline().decode('ascii')
    remove = list(map(float, featherData.split(' ')))
    for x in remove:
        withres = x*res
        data= np.append(data, withres)
        
    #Split data array into 4 seperate arrays based off what sensor the data comes from
    dataH0 = data[0::8]
    dataH1 = data[1::8]
    dataH2 = data[2::8]
    dataH3 = data[3::8]
    dataV0 = data[4::8]
    dataV1 = data[5::8]
    dataV2 = data[6::8]
    dataV3 = data[7::8]
   
    i = i+1
ser.close()

#Removing offset from each array of data

for index in range(len(dataH0)):
    if dataH0[index] >= avgH0:
        dataH0[index] = dataH0[index] - avgH0
    if dataH1[index] >= avgH1:
        dataH1[index] = dataH1[index] - avgH1
    if dataH2[index] >= avgH2:
        dataH2[index] = dataH2[index] - avgH2
    if dataH3[index] >= avgH3:
        dataH3[index] = dataH3[index] - avgH3
    if dataV0[index] >= avgV0:
        dataV0[index] = dataV0[index] - avgV0
    if dataV1[index] >= avgV1:
        dataV1[index] = dataV1[index] - avgV1
    if dataV2[index] >= avgV2:
        dataV2[index] = dataV2[index] - avgV2
    if dataV3[index] >= avgV3:
        dataV3[index] = dataV3[index] - avgV3
        

currentTime = nano()
dur = (currentTime - start)*1e-9

period = dur/end
freq = 1/period
xrange = np.arange(0,dur,period)

#filter

#cutfreq = 24

b,a = butter(2, 0.08)
outdataH0 = signal.lfilter(b,a, dataH0)
outdataH1 = signal.lfilter(b,a, dataH1)
outdataH2 = signal.lfilter(b,a, dataH2)
outdataH3 = signal.lfilter(b,a, dataH3)
outdataV0 = signal.lfilter(b,a, dataV0)
outdataV1 = signal.lfilter(b,a, dataV1)
outdataV2 = signal.lfilter(b,a, dataV2)
outdataV3 = signal.lfilter(b,a, dataV3)


#make voltage to newtons
forceH0 = 857*(outdataH0)+20.4
forceH1 = 304*(outdataH1)+11
forceH2 = 636*(outdataH2)+5.85
forceH3 = 198*(outdataH3)+15.3
forceV0 = 28081*(outdataV0)+6.75
forceV1 = 17498*(outdataV1)+14.2
forceV2 = 14072*(outdataV2)+11.3
forceV3 = 17651*(outdataV3)+19.9


#plot
plotforce = plt.figure(2)
plt.plot(xrange, forceH0,'xkcd:shocking pink', label = 'Force Horizontal A0')
plt.plot(xrange, forceH1,'xkcd:neon green', label = 'Force Horizontal A1')
plt.plot(xrange, forceH2,'xkcd:chocolate', label = 'Force Horizontal A2')
plt.plot(xrange, forceH3,'xkcd:tangerine', label = 'Force Horizontal A3')
plt.plot(xrange, forceV0,'xkcd:cherry red', label = 'Force Vertical A0')
plt.plot(xrange, forceV1,'xkcd:easter purple', label = 'Force Vertical A1')
plt.plot(xrange, forceV2,'xkcd:pure blue', label = 'Force Vertical A2')
plt.plot(xrange, forceV3,'xkcd:yellowish', label = 'Force Vertical A3')
plt.ylim(0)

plt.xlabel('Seconds')
plt.ylabel('Newtons')
plt.title('Force vs. Time')
plt.grid(color='k', linestyle='-', linewidth=0.2)
plt.legend()



#Print out Parameters
avgfv0 = offset(forceV0)
avgfv1 = offset(forceV1)
avgfv2 = offset(forceV2)
avgfv3 = offset(forceV3)
avgfh0 = offset(forceH0)
avgfh1 = offset(forceH1)
avgfh2 = offset(forceH2)
avgfh3 = offset(forceH3)
'''old x and y
forceoutX = (avgfh0+avgfh1+avgfh2+avgfh3)/4
forceoutY = (avgfv0+avgfv1+avgfv2+avgfv3)/4
'''
forceoutX = (avgfh1 + avgfh2)/2
forceoutZ = (avgfh0 + avgfh3)/2
forceoutY = (avgfv0+avgfv1+avgfv2+avgfv3)/4

maxV0 = np.max(forceV0)
maxV1 = np.max(forceV1)
maxV2 = np.max(forceV2)
maxV3 = np.max(forceV3)
if maxV0 > maxV1 and maxV0 > maxV2 and maxV0 > maxV3:
    peakForceY = maxV0
elif maxV1 > maxV0 and maxV1 > maxV2 and maxV1 > maxV3:
    peakForceY = maxV1
elif maxV2 > maxV0 and maxV2 > maxV1 and maxV2 > maxV3:
    peakForceY = maxV2
else:
    peakForceY = maxV3
maxH0 = np.max(forceH0)
maxH1 = np.max(forceH1)
maxH2 = np.max(forceH2)
maxH3 = np.max(forceH3)
'''old max horizontal
if maxH0 > maxH1 and maxH0 > maxH2 and maxH0 > maxH3:
    peakForceX = maxH0
elif maxH1 > maxH0 and maxH1 > maxH2 and maxH1 > maxH3:
    peakForceX = maxH1
elif maxH2 > maxH0 and maxH2 > maxH1 and maxH2 > maxH3:
    peakForceX = maxH2
else:
    peakForceX = maxH3
'''
if maxH1 > maxH2:
    peakForceX = maxH1
else:
    peakForceX = maxH2

if maxH0 > maxH3:
    peakForceZ = maxH0
else:
    peakForceZ = maxH3

forceat0 = kilo*gravity
netforceX = forceoutX - (kilo*gravity)
netforceY = forceoutY - (kilo*gravity)
netforceZ = forceoutZ - (kilo*gravity)
accelX = netforceX/kilo
accelY = netforceY/kilo
accelZ = netforceZ/kilo
ImpX = forceoutX * dur
ImpY = forceoutY * dur
ImpZ = forceoutZ * dur
#try to do velocity at just a certain point
velocityX = accelX * dur
velocityY = accelY * dur
velocityZ = accelZ * dur
peakPowerX = peakForceX * velocityX
peakPowerY = peakForceY * velocityY
peakPowerZ = peakForceZ * velocityZ
print('Downward Force at 0 m/s = ', forceat0, 'N')
print('Acceleration in the X direction = ', accelX,'m/s^2')
print('Acceleration in the Y direction = ', accelY,'m/s^2')
print('Acceleration in the Z direction = ', accelZ,'m/s^2')
print('Velocity in X direction = ', velocityX, 'm/s')
print('Velocity in Y direction = ', velocityY, 'm/s')
print('Velocity in Z direction = ', velocityZ, 'm/s')
print('Peak force in the X direction= ', peakForceX, 'N')
print('Peak force in the Y direction= ', peakForceY, 'N')
print('Peak force in the Z direction= ', peakForceZ, 'N')
print('Peak power in the X direction= ', peakPowerX, 'Nm/s')
print('Peak power in the Y direction= ', peakPowerY, 'Nm/s')
print('Peak power in the Z direction= ', peakPowerZ, 'Nm/s')
print('Impulse in the X direction = ', ImpX, 'Ns')
print('Impulse in the Y direction = ', ImpY, 'Ns')
print('Impulse in the Z direction = ', ImpZ, 'Ns')
print('Net force in the X direction= ', netforceX, 'N')
print('Net force in the Y direction= ', netforceY, 'N')
print('Net force in the Z direction= ', netforceZ, 'N')
print('Duration = ', dur, 's')

plt.show()

# Adding to sheet
bold = workbook.add_format({'bold': True})
worksheet.set_column(0, 1, 15)
worksheet.set_column(2, 25, 30)
worksheet.write('A1','Athlete:', bold)
worksheet.write('B1', name)
worksheet.write('A2','Weight (kg):', bold)
worksheet.write('B2', kilo)
worksheet.write('C1', 'Parameter', bold)
worksheet.write('D1', 'Measurement', bold)

calcs = (
    ['Downward Force at 0 m/s (N)', round(forceat0, 5)],
    ['Acceleration in the X direction (m/s^2)', round(accelX, 5)],
    ['Acceleration in the Y direction (m/s^2)', round(accelY, 5)],
    ['Acceleration in the Z direction (m/s^2)', round(accelZ, 5)],
    ['Velocity in X direction (m/s)', round(velocityX, 5)],
    ['Velocity in Y direction (m/s)', round(velocityY, 5)],
    ['Velocity in Z direction (m/s)', round(velocityZ, 5)],
    ['Peak force in the X direction (N)', round(peakForceX,5)],
    ['Peak force in the Y direction (N)', round(peakForceY,5)],
    ['Peak force in the Z direction (N)', round(peakForceZ,5)],
    ['Peak power in the X direction (Nm/s) ', round(peakPowerX,5)],
    ['Peak power in the Y direction (Nm/s) ', round(peakPowerY,5)],
    ['Peak power in the Z direction (Nm/s) ', round(peakPowerZ,5)],
    ['Impulse in the X direction (Ns) ', round(ImpX,5)],
    ['Impulse in the Y direction (Ns) ', round(ImpY,5)],
    ['Impulse in the Z direction (Ns) ', round(ImpZ,5)],
    ['Net force in the X direction (N) ', round(netforceX,5)],
    ['Net force in the Y direction (N) ', round(netforceY,5)],
    ['Net force in the Z direction (N) ', round(netforceZ,5)],
    )
row = 1
col = 2

for parameter, measurement in (calcs):
    worksheet.write(row, col,   parameter)
    worksheet.write(row, col + 1, measurement)
    row+=1
      
#plot on excel sheet

worksheet.write('E1','Time')
worksheet.write_column(1, 4, xrange)
force_lst = {'ForceH0': forceH0,
             'ForceH1': forceH1,
             'ForceH2': forceH2,
             'ForceH3': forceH3,
             'ForceV0': forceV0,
             'ForceV1': forceV1,
             'ForceV2': forceV2,
             'ForceV3': forceV3,
             }
col_num = 5
for labels, value in force_lst.items():
    worksheet.write(0, col_num, labels)
    worksheet.write_column(1, col_num, value)
    col_num += 1
    
chart = workbook.add_chart({'type': 'line'})

chart.add_series({'name':'ForceH0',
                  'categories':'=Sheet1!$E$2:$E$1001',
                  'values':'=Sheet1!$F$2:$F$1001',
                  })
chart.add_series({'name':'ForceH1',
                  'categories':'=Sheet1!$E$2:$E$1001',
                  'values':'=Sheet1!$G$2:$G$1001',
                  })
chart.add_series({'name':'ForceH2',
                  'categories':'=Sheet1!$E$2:$E$1001',
                  'values':'=Sheet1!$H$2:$H$1001',
                  })
chart.add_series({'name':'ForceH3',
                  'categories':'=Sheet1!$E$2:$E$1001',
                  'values':'=Sheet1!$I$2:$I$1001',
                  })
chart.add_series({'name':'ForceV0',
                  'categories':'=Sheet1!$E$2:$E$1001',
                  'values':'=Sheet1!$J$2:$J$1001',
                  })
chart.add_series({'name':'ForceV1',
                  'categories':'=Sheet1!$E$2:$E$1001',
                  'values':'=Sheet1!$K$2:$K$1001',
                  })
chart.add_series({'name':'ForceV2',
                  'categories':'=Sheet1!$E$2:$E$1001',
                  'values':'=Sheet1!$L$2:$L$1001',
                  })
chart.add_series({'name':'ForceV3',
                  'categories':'=Sheet1!$E$2:$E$1001',
                  'values':'=Sheet1!$M$2:$M$1001',
                  })

chart.set_title ({'name': 'Force over Time'})
chart.set_x_axis({'name': 'Time (s)'})
chart.set_y_axis({'name': 'Force (N)'})

worksheet.insert_chart('O8', chart)


workbook.close()



