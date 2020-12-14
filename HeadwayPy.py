# -*- coding: utf-8 -*-
"""
Created on Tue Nov 10 12:19:52 2020

@author: tfahry
"""


#%%init (when signals are changed) 
###############never ever perform Excel writes in this cell!!!!!!!!!!
import xlwings as xw


def unpop(cells):
    output = []
    for i in cells:
        if i is not None:
            output.append(i)
    return output

master = xw.Book('Headways from Headways CSV test mergebook Quickened 6.xlsm')
signalsheet  = master.sheets["Signals CSV"]
intersheet  = master.sheets["Interlockings CSV"]
trainrunsheet = master.sheets["Train run"]

signumbers = unpop(signalsheet.range('D2:D9999').value)
siglines = unpop(signalsheet.range('N2:N9999').value)
sigkeys = ['{}{}'.format(line,number) for number, line in zip(signumbers,siglines)]

siginterlockingtype = unpop(signalsheet.range('F2:F9999').value)
sigdict = dict(zip(sigkeys,siginterlockingtype))

internames = unpop(intersheet.range('B2:B9999').value) #interlocking name
intertimes = unpop(intersheet.range('C2:C9999').value) #setting time from Interlockings CSV
interdict = dict(zip(internames,intertimes))

trainrunnumbers = unpop(trainrunsheet.range('F2:F9999').value)
trainrunlines = [signal.split('/')[-1] for signal in unpop(trainrunsheet.range('D2:D9999').value)]
trainrunkeys = ['{}{}'.format(line,number) for number, line in zip(trainrunnumbers,trainrunlines)]

trainrunsysinfo = unpop(trainrunsheet.range('K2:K9999').value)
trainrundict = dict(zip(trainrunkeys,trainrunsysinfo))

titlerow = ['Station',
 'Station name',
 'Line1',
 'Track1',
 'Name signal1',
 'Signal1',
 'Km1',
 'Line2',
 'Track2',
 'Name signal2',
 'Signal2',
 'Km2',
 'Length',
 'BloOccTime',
 'StartOcc',
 'Arrival',
 'RelSig',
 'Overlap cleared (RailSys)',
 'PrecTr',
 'HeadwayPrec',
 'BufferTPrec',
 'HeadwPrecSig',
 'FolTr',
 'HeadwayFol',
 'BufferTFol',
 'HeadwFolSig']

#%%run
wb = xw.Book('Headways from Headways CSV test mergebook Quickened 6.xlsm')
sheet  = wb.sheets["Headways CSV"]

# if sheet.range('A1').value != 'Station':
#     sheet.range("1:1").api.Insert()
#     sheet.range('A1').value = titlerow

######reads
signal2  = unpop(sheet.range('M2:M999').value)  #'Number' in signals csv
line2 = unpop(sheet.range('J2:J999').value)     #'Line/global area' in signals csv
startocc = unpop(sheet.range('Q2:Q999').value)
             
##########processing
line2 = [line.split('/')[-1] for line in line2]
signal2 = ['{}{}'.format(line,number) for line, number in zip(line2,signal2)]

sigtypes = [sigdict[signal] for signal in signal2] #interlocking type

sysinfo = [trainrundict[signal] for signal in signal2] #TrainConSysInfo

    

#ofot = unpop(sheet.range('C3:C999').value)

aaso = []
aast = []
offsets = []

offset_minus = 3

for i, time in enumerate(startocc):
    try:
        offset = int(sigdict[signal2[i]][0])               #get aspect number of signal
        st = interdict[sigdict[signal2[i+offset-offset_minus]]]/86400 #get setting time of interlocking
                      
        aaso.append(startocc[i+offset-offset_minus]+st)
        offsets.append(offset)
        aast.append(st)
        
    except IndexError:
        aaso.append(None)
        offsets.append(None)
        aast.append(None)


##%%writes
wb = xw.Book('Headways from Headways CSV test mergebook Quickened 6.xlsm')
sheet  = wb.sheets["Headways CSV"]

sheet.range('X1').value = 'TrainConSysInfo'
sheet.range('X2').value = list(zip(sysinfo))

sheet.range('U1').value = 'StartOcc[AA-{}]'.format(offset_minus)
sheet.range('U2').value = list(zip(aaso[1:]))

# sheet.range('V1').value = 'SetTime[AA]'
# sheet.range('V2').value = list(zip(aast))
        
# sheet.range('W1').value = 'Signal2 Interlocking type'
# sheet.range('W2').value = list(zip(sigtypes))

# sheet.range('X1').value = 'Headways'
# sheet.range('X2:X999').value = '=@INDIRECT("R["&RC26-1&"]C20",FALSE)-@INDIRECT("R["&RC26-2&"]C17",FALSE)-RC22'

# sheet.range('Y1').value = 'Aspect No.'
# sheet.range('Y2').value = list(zip(offsets))

#=@INDIRECT("R["&RC26-1&"]C20",FALSE)-@INDIRECT("R["&RC26-2&"]C17",FALSE)-RC22











#%%