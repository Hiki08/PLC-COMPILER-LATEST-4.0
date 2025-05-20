# %%
#Load PyFiles
from Imports import *
import DateAndTimeManager

from FilesReader import *

from Em2p import em2P
from Em3p import em3P
from Fm import fM
from Dfb import dFB
from Dfb import Tensile
from Rdb import rDB
from Csb import cSB

#%%
# num = [2, 5, 3, "nan"]

# print(statistics.mean(num))

#%%
# DateAndTimeManager.GetDateToday()

# filesreader = filesReader()
# filesreader.readingYearStored = DateAndTimeManager.yearNow
# filesreader.ReadEm2pFiles()
# filesreader.ReadEm3pFiles()
# filesreader.ReadFmFiles()


# filesreader.ReadDfbSnapFiles()
# filesreader.ReadDfbFiles()
# filesreader.ReadTensile()

# filesreader.ReadRdbCheckSheetFiles()
# filesreader.ReadRdbFiles()

# filesreader.ReadCsbFiles()

# em2p = em2P()
# em2p.GettingData("EM0580106P", "CAT-4J15DI")

# em2p.GettingData("EM0660046P", "FC6030-1F18GT")
# em2p.GettingData("EM0580106P", "CAT-4J15DI")

# em3p = em3P()
# em3p.GettingData("EM0580107P", "CAT-4J15DI")
# em3p.GettingData("EM0580107P", "CAT-4I14DI")

# fm = fM()
# fm.GettingData("FM05000102-00A", "112524A-40")

# dfb = dFB()
# dfb.ReadDfbSnap("20241031-A")
# dfb.GettingData("DFB6600600")

# tensile = Tensile()
# tensile.GettingData("DFB6600600", dfb.dfbLotNumber2[:-3])

# rdb = rDB()
# rdb.ReadCheckSheet("20241209-D", "RDB5200200")
# rdb.GettingData("RDB5200200")

# csb = cSB()
# csb.GettingData("CSB6400802", "012225A")



# %%
#Variables 
dfVt1 = ""
dfVt2 = ""
dfVt3 = ""
dfVt4 = ""
dfVt5 = ""
dfVt6 = ""

tempDfVt1 = ""
tempDfVt2 = ""
tempDfVt3 = ""
tempDfVt4 = ""
tempDfVt5 = ""
tempDfVt6 = ""

process1Status = ""
process2Status = ""
process3Status = ""
process4Status = ""
process5Status = ""
process6Status = ""
isRepairedWithNG = False
piStatus = ""

compiledFrame = ""
excelData = ""
ngProcess = ""
repairedProcess = ""

compiledFrame2 = ""
excelData2 = ""
processData = ""

process1Row = 0
process2Row = 0
process3Row = 0
process4Row = 0
process5Row = 0
process6Row = 0
piRow = 0

processPendingToRepair = []

canCompile = False

canCompilePI = False

dfPi = ""
dfPiNotDone = ""
tempdfPi = ""

isCsvReaded = False

readCount = 0

programRunning = True

dateToRead = "2024/11/05"
dateToReadDashFormat = "2024-11-05"

#UI Variables

compileButton = ""
autoRunButton = ""
autoRun = False

loadingText = "Loading"

time_picker = ""

frame1 = ""
frame2 = ""

calendarPicker = ""

# %%
def CheckPICsv():
    global dfPi
    global canCompilePI
    global dfPiNotDone

    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)

    canCompilePI = False
    
    piDirectory = (r'\\192.168.2.19\ai_team\AI Program\Outputs\PICompiled')
    os.chdir(piDirectory)

    dfPi = pd.read_csv(f'PICompiled{dateToReadDashFormat}.csv', encoding='latin1')
    
    dfPiNotDone = dfPi[(dfPi["CHECKING"].isin(["-"])) & (dfPi["PROCESS S/N"].isin(["MASTER PUMP"])) | (dfPi["PROCESS S/N"].isin(["RUNNING"]))]
    if len(dfPiNotDone) != 0:
        canCompilePI = True
    else:
        canCompilePI = False

# %%
def CompilePICsv():
    global dfPi
    global dfPiNotDone
    global tempdfPi
    global canCompilePI
    global compiledFrame

    global piRow

    for a in range(0, len(dfPiNotDone)):
        piRow += 1

        tempdfPi = dfPiNotDone.iloc[[a], :]

        if tempdfPi["PROCESS S/N"].values[0] == "MASTER PUMP":
            processData = "MASTER PUMP"
        elif tempdfPi["PROCESS S/N"].values[0] == "RUNNING":
            processData = "RUNNING"

        # piDirectory = (r'\\192.168.2.19\ai_team\AI Program\Outputs')
        # os.chdir(piDirectory)

        # tempdfPi = dfPi.iloc[[a], :]
        # dfPi.loc[dfPi["TIME"] == tempdfPi["TIME"].values[0], "CHECKING"] = "Done"
        # dfPi.to_csv(f"PICompiled.csv", index = False)
        

        excelData2 = {
                    "DATE": tempdfPi["DATE"].values,
                    "TIME": tempdfPi["TIME"].values,
                    "MODEL CODE": tempdfPi["MODEL CODE"].values,
                    "PROCESS S/N": tempdfPi["PROCESS S/N"].values,
                    "S/N": tempdfPi["S/N"].values,
                    "PASS/NG": tempdfPi["PASS/NG"].values,
                    "VOLTAGE MAX (V)": tempdfPi["VOLTAGE MAX (V)"].values,
                    "WATTAGE MAX (W)": tempdfPi["WATTAGE MAX (W)"].values,
                    "CLOSED PRESSURE_MAX (kPa)": tempdfPi["CLOSED PRESSURE_MAX (kPa)"].values,
                    "VOLTAGE Middle (V)": tempdfPi["VOLTAGE Middle (V)"].values,
                    "WATTAGE Middle (W)": tempdfPi["WATTAGE Middle (W)"].values,
                    "AMPERAGE Middle (A)": tempdfPi["AMPERAGE Middle (A)"].values,
                    "CLOSED PRESSURE Middle (kPa)": tempdfPi["CLOSED PRESSURE Middle (kPa)"].values,
                    "dB(A) 1": tempdfPi["dB(A) 1"].values,
                    "dB(A) 2": tempdfPi["dB(A) 2"].values,
                    "dB(A) 3": tempdfPi["dB(A) 3"].values,
                    "VOLTAGE MIN (V)": tempdfPi["VOLTAGE MIN (V)"].values,
                    "WATTAGE MIN (W)": tempdfPi["WATTAGE MIN (W)"].values,
                    "CLOSED PRESSURE MIN (kPa)": tempdfPi["CLOSED PRESSURE MIN (kPa)"].values,
            
                    "Process 1 Model Code" : [processData],
                    "Process 1 S/N" : [processData],
                    "Process 1 ID" : [processData],
                    "Process 1 NAME" : [processData],
                    "Process 1 Regular/Contractual" : [processData],
                    "Process 1 Em2p" : [processData],
                    "Process 1 Em2p Lot No" : [processData],
                    "Process 1 Em2p Inspection 3 Average Data" : [processData],
                    "Process 1 Em2p Inspection 4 Average Data" : [processData],
                    "Process 1 Em2p Inspection 5 Average Data" : [processData],
                    "Process 1 Em2p Inspection 10 Average Data" : [processData],
                    "Process 1 Em2p Inspection 3 Minimum Data" : [processData],
                    "Process 1 Em2p Inspection 4 Minimum Data" : [processData],
                    "Process 1 Em2p Inspection 5 Minimum Data" : [processData],
                    "Process 1 Em2p Inspection 3 Maximum Data" : [processData],
                    "Process 1 Em2p Inspection 4 Maximum Data" : [processData],
                    "Process 1 Em2p Inspection 5 Maximum Data" : [processData],
                    "Process 1 Em3p" : [processData],
                    "Process 1 Em3p Lot No" : [processData],
                    "Process 1 Em3p Inspection 3 Average Data" : [processData],
                    "Process 1 Em3p Inspection 4 Average Data" : [processData],
                    "Process 1 Em3p Inspection 5 Average Data" : [processData],
                    "Process 1 Em3p Inspection 10 Average Data" : [processData],
                    "Process 1 Em3p Inspection 3 Minimum Data" : [processData],
                    "Process 1 Em3p Inspection 4 Minimum Data" : [processData],
                    "Process 1 Em3p Inspection 5 Minimum Data" : [processData],
                    "Process 1 Em3p Inspection 3 Maximum Data" : [processData],
                    "Process 1 Em3p Inspection 4 Maximum Data" : [processData],
                    "Process 1 Em3p Inspection 5 Maximum Data" : [processData],
                    "Process 1 Harness" : [processData],
                    "Process 1 Harness Lot No" : [processData],
                    "Process 1 Frame" : [processData],
                    "Process 1 Frame Lot No" : [processData],
                    "Process 1 Frame Inspection 1 Average Data" : [processData], 
                    "Process 1 Frame Inspection 2 Average Data" : [processData], 
                    "Process 1 Frame Inspection 3 Average Data" : [processData], 
                    "Process 1 Frame Inspection 4 Average Data" : [processData], 
                    "Process 1 Frame Inspection 5 Average Data" : [processData], 
                    "Process 1 Frame Inspection 6 Average Data" : [processData], 
                    "Process 1 Frame Inspection 7 Average Data" : [processData], 
                    "Process 1 Frame Inspection 1 Minimum Data" : [processData], 
                    "Process 1 Frame Inspection 2 Minimum Data" : [processData], 
                    "Process 1 Frame Inspection 3 Minimum Data" : [processData], 
                    "Process 1 Frame Inspection 4 Minimum Data" : [processData], 
                    "Process 1 Frame Inspection 5 Minimum Data" : [processData], 
                    "Process 1 Frame Inspection 6 Minimum Data" : [processData], 
                    "Process 1 Frame Inspection 7 Minimum Data" : [processData], 
                    "Process 1 Frame Inspection 1 Maximum Data" : [processData], 
                    "Process 1 Frame Inspection 2 Maximum Data" : [processData], 
                    "Process 1 Frame Inspection 3 Maximum Data" : [processData], 
                    "Process 1 Frame Inspection 4 Maximum Data" : [processData], 
                    "Process 1 Frame Inspection 5 Maximum Data" : [processData], 
                    "Process 1 Frame Inspection 6 Maximum Data" : [processData], 
                    "Process 1 Frame Inspection 7 Maximum Data" : [processData], 
                    "Process 1 Bushing" : [processData],
                    "Process 1 Bushing Lot No" : [processData],
                    "Process 1 ST" : [processData],
                    "Process 1 Actual Time" : [processData],
                    "Process 1 NG Cause" : [processData],
                    "Process 1 Repaired Action" : [processData],

                    "Process 2 Model Code" : [processData],
                    "Process 2 S/N" : [processData],
                    "Process 2 ID" : [processData],
                    "Process 2 NAME" : [processData],
                    "Process 2 Regular/Contractual" : [processData],
                    "Process 2 M4x40 Screw" : [processData],
                    "Process 2 M4x40 Screw Lot No" : [processData],
                    "Process 2 Rod Blk" : [processData],
                    "Process 2 Rod Blk Lot No" : [processData],
                    "Process 2 Rod Blk Tesla 1 Average Data" : [processData],
                    "Process 2 Rod Blk Tesla 2 Average Data" : [processData],
                    "Process 2 Rod Blk Tesla 3 Average Data" : [processData],
                    "Process 2 Rod Blk Tesla 4 Average Data" : [processData],
                    "Process 2 Rod Blk Tesla 1 Minimum Data" : [processData],
                    "Process 2 Rod Blk Tesla 2 Minimum Data" : [processData],
                    "Process 2 Rod Blk Tesla 3 Minimum Data" : [processData],
                    "Process 2 Rod Blk Tesla 4 Minimum Data" : [processData],
                    "Process 2 Rod Blk Tesla 1 Maximum Data" : [processData],
                    "Process 2 Rod Blk Tesla 2 Maximum Data" : [processData],
                    "Process 2 Rod Blk Tesla 3 Maximum Data" : [processData],
                    "Process 2 Rod Blk Tesla 4 Maximum Data" : [processData],
                    "Process 2 Rod Blk Inspection 1 Average Data" : [processData],
                    "Process 2 Rod Blk Inspection 2 Average Data" : [processData],
                    "Process 2 Rod Blk Inspection 3 Average Data" : [processData],
                    "Process 2 Rod Blk Inspection 4 Average Data" : [processData],
                    "Process 2 Rod Blk Inspection 5 Average Data" : [processData],
                    "Process 2 Rod Blk Inspection 6 Average Data" : [processData],
                    "Process 2 Rod Blk Inspection 7 Average Data" : [processData],
                    "Process 2 Rod Blk Inspection 8 Average Data" : [processData],
                    "Process 2 Rod Blk Inspection 9 Average Data" : [processData],
                    "Process 2 Rod Blk Inspection 1 Minimum Data" : [processData],
                    "Process 2 Rod Blk Inspection 2 Minimum Data" : [processData],
                    "Process 2 Rod Blk Inspection 3 Minimum Data" : [processData],
                    "Process 2 Rod Blk Inspection 4 Minimum Data" : [processData],
                    "Process 2 Rod Blk Inspection 5 Minimum Data" : [processData],
                    "Process 2 Rod Blk Inspection 6 Minimum Data" : [processData],
                    "Process 2 Rod Blk Inspection 7 Minimum Data" : [processData],
                    "Process 2 Rod Blk Inspection 8 Minimum Data" : [processData],
                    "Process 2 Rod Blk Inspection 9 Minimum Data" : [processData],
                    "Process 2 Rod Blk Inspection 1 Maximum Data" : [processData],
                    "Process 2 Rod Blk Inspection 2 Maximum Data" : [processData],
                    "Process 2 Rod Blk Inspection 3 Maximum Data" : [processData],
                    "Process 2 Rod Blk Inspection 4 Maximum Data" : [processData],
                    "Process 2 Rod Blk Inspection 5 Maximum Data" : [processData],
                    "Process 2 Rod Blk Inspection 6 Maximum Data" : [processData],
                    "Process 2 Rod Blk Inspection 7 Maximum Data" : [processData],
                    "Process 2 Rod Blk Inspection 8 Maximum Data" : [processData],
                    "Process 2 Rod Blk Inspection 9 Maximum Data" : [processData],
                    "Process 2 Df Blk" : [processData],
                    "Process 2 Df Blk Lot No" : [processData],
                    "Process 2 Df Blk Inspection 1 Average Data" : [processData],
                    "Process 2 Df Blk Inspection 2 Average Data" : [processData],
                    "Process 2 Df Blk Inspection 3 Average Data" : [processData],
                    "Process 2 Df Blk Inspection 4 Average Data" : [processData],
                    "Process 2 Df Blk Inspection 1 Minimum Data" : [processData],
                    "Process 2 Df Blk Inspection 2 Minimum Data" : [processData],
                    "Process 2 Df Blk Inspection 3 Minimum Data" : [processData],
                    "Process 2 Df Blk Inspection 4 Minimum Data" : [processData],
                    "Process 2 Df Blk Inspection 1 Maximum Data" : [processData],
                    "Process 2 Df Blk Inspection 2 Maximum Data" : [processData],
                    "Process 2 Df Blk Inspection 3 Maximum Data" : [processData],
                    "Process 2 Df Blk Inspection 4 Maximum Data" : [processData],
                    "Process 2 Df Blk Tensile Rate Of Change Average" : [processData],
                    "Process 2 Df Blk Tensile Rate Of Change Minimum" : [processData],
                    "Process 2 Df Blk Tensile Rate Of Change Maximum" : [processData],
                    "Process 2 Df Blk Tensile Start Force Average" : [processData],
                    "Process 2 Df Blk Tensile Start Force Minimum" : [processData],
                    "Process 2 Df Blk Tensile Start Force Maximum" : [processData],
                    "Process 2 Df Blk Tensile Terminating Force Average" : [processData],
                    "Process 2 Df Blk Tensile Terminating Force Minimum" : [processData],
                    "Process 2 Df Blk Tensile Terminating Force Maximum" : [processData],
                    "Process 2 Df Ring" : [processData],
                    "Process 2 Df Ring Lot No" : [processData],
                    "Process 2 Washer" : [processData],
                    "Process 2 Washer Lot No" : [processData],
                    "Process 2 Lock Nut" : [processData],
                    "Process 2 Lock Nut Lot No" : [processData],
                    "Process 2 ST" : [processData],
                    "Process 2 Actual Time" : [processData],
                    "Process 2 NG Cause" : [processData],
                    "Process 2 Repaired Action" : [processData],

                    "Process 3 Model Code" : [processData],
                    "Process 3 S/N" : [processData],
                    "Process 3 ID" : [processData],
                    "Process 3 NAME" : [processData],
                    "Process 3 Regular/Contractual" : [processData],
                    "Process 3 Frame Gasket" : [processData],
                    "Process 3 Frame Gasket Lot No" : [processData],
                    "Process 3 Casing Block" : [processData],
                    "Process 3 Casing Block Lot No" : [processData],
                    "Process 3 Casing Block Inspection 1 Average Data" : [processData],
                    "Process 3 Casing Block Inspection 1 Minimum Data" : [processData],
                    "Process 3 Casing Block Inspection 1 Maximum Data" : [processData],
                    "Process 3 Casing Gasket" : [processData],
                    "Process 3 Casing Gasket Lot No" : [processData],
                    "Process 3 M4x16 Screw 1" : [processData],
                    "Process 3 M4x16 Screw 1 Lot No" : [processData],
                    "Process 3 M4x16 Screw 2" : [processData],
                    "Process 3 M4x16 Screw 2 Lot No" : [processData],
                    "Process 3 Ball Cushion" : [processData],
                    "Process 3 Ball Cushion Lot No" : [processData],
                    "Process 3 Frame Cover" : [processData],
                    "Process 3 Frame Cover Lot No" : [processData],
                    "Process 3 Partition Board" : [processData],
                    "Process 3 Partition Board Lot No" : [processData],
                    "Process 3 Built In Tube 1" : [processData],
                    "Process 3 Built In Tube 1 Lot No" : [processData],
                    "Process 3 Built In Tube 2" : [processData],
                    "Process 3 Built In Tube 2 Lot No" : [processData],
                    "Process 3 Head Cover" : [processData],
                    "Process 3 Head Cover Lot No" : [processData],
                    "Process 3 Casing Packing" : [processData],
                    "Process 3 Casing Packing Lot No" : [processData],
                    "Process 3 M4x12 Screw" : [processData],
                    "Process 3 M4x12 Screw Lot No" : [processData],
                    "Process 3 Csb L" : [processData],
                    "Process 3 Csb L Lot No" : [processData],
                    "Process 3 Csb R" : [processData],
                    "Process 3 Csb R Lot No" : [processData],
                    "Process 3 Head Packing" : [processData],
                    "Process 3 Head Packing Lot No" : [processData],
                    "Process 3 ST" : [processData],
                    "Process 3 Actual Time" : [processData],
                    "Process 3 NG Cause" : [processData],
                    "Process 3 Repaired Action" : [processData],

                    "Process 4 Model Code" : [processData],
                    "Process 4 S/N" : [processData],
                    "Process 4 ID" : [processData],
                    "Process 4 NAME" : [processData],
                    "Process 4 Regular/Contractual" : [processData],
                    "Process 4 Tank" : [processData],
                    "Process 4 Tank Lot No" : [processData],
                    "Process 4 Upper Housing" : [processData],
                    "Process 4 Upper Housing Lot No" : [processData],
                    "Process 4 Cord Hook" : [processData],
                    "Process 4 Cord Hook Lot No" : [processData],
                    "Process 4 M4x16 Screw" : [processData],
                    "Process 4 M4x16 Screw Lot No" : [processData],
                    "Process 4 Tank Gasket" : [processData],
                    "Process 4 Tank Gasket Lot No" : [processData],
                    "Process 4 Tank Cover" : [processData],
                    "Process 4 Tank Cover Lot No" : [processData],
                    "Process 4 Housing Gasket" : [processData],
                    "Process 4 Housing Gasket Lot No" : [processData],
                    "Process 4 M4x40 Screw" : [processData],
                    "Process 4 M4x40 Screw Lot No" : [processData],
                    "Process 4 PartitionGasket" : [processData],
                    "Process 4 PartitionGasket Lot No" : [processData],
                    "Process 4 M4x12 Screw" : [processData],
                    "Process 4 M4x12 Screw Lot No" : [processData],
                    "Process 4 Muffler" : [processData],
                    "Process 4 Muffler Lot No" : [processData],
                    "Process 4 Muffler Gasket" : [processData],
                    "Process 4 Muffler Gasket Lot No" : [processData],
                    "Process 4 VCR" : [processData],
                    "Process 4 VCR Lot No" : [processData],
                    "Process 4 ST" : [processData],
                    "Process 4 Actual Time" : [processData],
                    "Process 4 NG Cause" : [processData],
                    "Process 4 Repaired Action" : [processData],
                    
                    "Process 5 Model Code" : [processData],
                    "Process 5 S/N" : [processData],
                    "Process 5 ID" : [processData],
                    "Process 5 NAME" : [processData],
                    "Process 5 Regular/Contractual" : [processData],
                    "Process 5 Rating Label" : [processData],
                    "Process 5 Rating Label Lot No" : [processData],
                    "Process 5 ST" : [processData],
                    "Process 5 Actual Time" : [processData],
                    "Process 5 NG Cause" : [processData],
                    "Process 5 Repaired Action" : [processData],
                    
                    "Process 6 Model Code" : [processData],
                    "Process 6 S/N" : [processData],
                    "Process 6 ID" : [processData],
                    "Process 6 NAME" : [processData],
                    "Process 6 Regular/Contractual" : [processData],
                    "Process 6 Vinyl" : [processData],
                    "Process 6 Vinyl Lot No" : [processData],
                    "Process 6 ST" : [processData],
                    "Process 6 Actual Time" : [processData],
                    "Process 6 NG Cause" : [processData],
                    "Process 6 Repaired Action" : [processData]
                }
        excelData2 = pd.DataFrame(excelData2)
        compiledFrame = pd.concat([compiledFrame, excelData2], ignore_index=True)

    canCompilePI = False

# %%
def WriteCsv(data):
    global dateToReadDashFormat

    fileDirectory = (r'\\192.168.2.19\ai_team\AI Program\Outputs\CompiledProcess')
    os.chdir(fileDirectory)
    print(os.getcwd())

    # if os.path.exists(f"{fileDirectory}/CompiledProcess{dateToReadDashFormat}.csv"):
    #     print("Overiting The Existing File")
    #     #Read Existed
    #     existedExcel = pd.read_csv(f"CompiledProcess{dateToReadDashFormat}.csv")
    #     newValue = pd.concat([existedExcel, data], axis = 0, ignore_index = True)
    #     wireFrame = newValue
    #     wireFrame.to_csv(f"CompiledProcess{dateToReadDashFormat}.csv", index = False)
    # else:
    
    print("Creating New File")
    newValue = pd.concat([data], axis = 0, ignore_index = True)
    wireFrame = newValue
    wireFrame.to_csv(f"CompiledProcess{dateToReadDashFormat}.csv", index = False)

# %%
def ReadCsv():
    global dfVt1
    global dfVt2
    global dfVt3
    global dfVt4
    global dfVt5
    global dfVt6

    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)

    vt1Directory = (r'\\192.168.2.19\ai_team\AI Program\Outputs\FC1 CSV\VT1')
    os.chdir(vt1Directory)
    dfVt1 = pd.read_csv('log000_1.csv', encoding='latin1')
    dfVt1.columns = ["Process 1 DATA No",
        "Process 1 DATE",
        "Process 1 TIME",
        "Process 1 Model Code",
        "Process 1 S/N",
        "Process 1 ID",
        "Process 1 NAME",
        "Process 1 Regular/Contractual",
        "Process 1 Em2p",
        "Process 1 Em2p Lot No",
        "Process 1 Em3p",
        "Process 1 Em3p Lot No",
        "Process 1 Harness",
        "Process 1 Harness Lot No",
        "Process 1 Frame",
        "Process 1 Frame Lot No",
        "Process 1 Bushing",
        "Process 1 Bushing Lot No",
        "Process 1 ST",
        "Process 1 Actual Time",
        "Process 1 NG Cause",
        "Process 1 Repaired Action"]
    
    vt2Directory = (r'\\192.168.2.19\ai_team\AI Program\Outputs\FC1 CSV\VT2')
    os.chdir(vt2Directory)
    dfVt2 = pd.read_csv('log000_2.csv', encoding='latin1')
    dfVt2.columns = ["Process 2 DATA No",
        "Process 2 DATE",
        "Process 2 TIME",
        "Process 2 Model Code",
        "Process 2 S/N",
        "Process 2 ID",
        "Process 2 NAME",
        "Process 2 Regular/Contractual",
        "Process 2 M4x40 Screw",
        "Process 2 M4x40 Screw Lot No",
        "Process 2 Rod Blk",
        "Process 2 Rod Blk Lot No",
        "Process 2 Df Blk",
        "Process 2 Df Blk Lot No",
        "Process 2 Df Ring",
        "Process 2 Df Ring Lot No",
        "Process 2 Washer",
        "Process 2 Washer Lot No",
        "Process 2 Lock Nut",
        "Process 2 Lock Nut Lot No",
        "Process 2 ST",
        "Process 2 Actual Time",
        "Process 2 NG Cause",
        "Process 2 Repaired Action"]

    vt3Directory = (r'\\192.168.2.19\ai_team\AI Program\Outputs\FC1 CSV\VT3')
    os.chdir(vt3Directory)
    dfVt3 = pd.read_csv('log000_3.csv', encoding='latin1')
    dfVt3.columns = ["Process 3 DATA No",
        "Process 3 DATE",
        "Process 3 TIME",
        "Process 3 Model Code",
        "Process 3 S/N",
        "Process 3 ID",
        "Process 3 NAME",
        "Process 3 Regular/Contractual",
        "Process 3 Frame Gasket",
        "Process 3 Frame Gasket Lot No",
        "Process 3 Casing Block",
        "Process 3 Casing Block Lot No",
        "Process 3 Casing Gasket",
        "Process 3 Casing Gasket Lot No",
        "Process 3 M4x16 Screw 1",
        "Process 3 M4x16 Screw 1 Lot No",
        "Process 3 M4x16 Screw 2",
        "Process 3 M4x16 Screw 2 Lot No",
        "Process 3 Ball Cushion",
        "Process 3 Ball Cushion Lot No",
        "Process 3 Frame Cover",
        "Process 3 Frame Cover Lot No",
        "Process 3 Partition Board",
        "Process 3 Partition Board Lot No",
        "Process 3 Built In Tube 1",
        "Process 3 Built In Tube 1 Lot No",
        "Process 3 Built In Tube 2",
        "Process 3 Built In Tube 2 Lot No",
        "Process 3 Head Cover",
        "Process 3 Head Cover Lot No",
        "Process 3 Casing Packing",
        "Process 3 Casing Packing Lot No",
        "Process 3 M4x12 Screw",
        "Process 3 M4x12 Screw Lot No",
        "Process 3 Csb L",
        "Process 3 Csb L Lot No",
        "Process 3 Csb R",
        "Process 3 Csb R Lot No",
        "Process 3 Head Packing",
        "Process 3 Head Packing Lot No",
        "Process 3 ST",
        "Process 3 Actual Time",
        "Process 3 NG Cause",
        "Process 3 Repaired Action"]

    vt4Directory = (r'\\192.168.2.19\ai_team\AI Program\Outputs\FC1 CSV\VT4')
    os.chdir(vt4Directory)
    dfVt4 = pd.read_csv('log000_4.csv', encoding='latin1')
    dfVt4.columns = ["Process 4 DATA No",
        "Process 4 DATE",
        "Process 4 TIME",
        "Process 4 Model Code",
        "Process 4 S/N",
        "Process 4 ID",
        "Process 4 NAME",
        "Process 4 Regular/Contractual",
        "Process 4 Tank",
        "Process 4 Tank Lot No",
        "Process 4 Upper Housing",
        "Process 4 Upper Housing Lot No",
        "Process 4 Cord Hook",
        "Process 4 Cord Hook Lot No",
        "Process 4 M4x16 Screw",
        "Process 4 M4x16 Screw Lot No",
        "Process 4 Tank Gasket",
        "Process 4 Tank Gasket Lot No",
        "Process 4 Tank Cover",
        "Process 4 Tank Cover Lot No",
        "Process 4 Housing Gasket",
        "Process 4 Housing Gasket Lot No",
        "Process 4 M4x40 Screw",
        "Process 4 M4x40 Screw Lot No",
        "Process 4 PartitionGasket",
        "Process 4 PartitionGasket Lot No",
        "Process 4 M4x12 Screw",
        "Process 4 M4x12 Screw Lot No",
        "Process 4 Muffler",
        "Process 4 Muffler Lot No",
        "Process 4 Muffler Gasket",
        "Process 4 Muffler Gasket Lot No",
        "Process 4 VCR",
        "Process 4 VCR Lot No",
        "Process 4 ST",
        "Process 4 Actual Time",
        "Process 4 NG Cause",
        "Process 4 Repaired Action"]

    vt5Directory = (r'\\192.168.2.19\ai_team\AI Program\Outputs\FC1 CSV\VT5')
    os.chdir(vt5Directory)
    dfVt5 = pd.read_csv('log000_5.csv', encoding='latin1')
    dfVt5.columns = ["Process 5 DATA No",
        "Process 5 DATE",
        "Process 5 TIME",
        "Process 5 Model Code",
        "Process 5 S/N",
        "Process 5 ID",
        "Process 5 NAME",
        "Process 5 Regular/Contractual",
        "Process 5 Rating Label",
        "Process 5 Rating Label Lot No",
        "Process 5 ST",
        "Process 5 Actual Time",
        "Process 5 NG Cause",
        "Process 5 Repaired Action"]

    vt6Directory = (r'\\192.168.2.19\ai_team\AI Program\Outputs\FC1 CSV\VT6')
    os.chdir(vt6Directory)
    dfVt6 = pd.read_csv('log000_6.csv', encoding='latin1')
    dfVt6.columns = ["Process 6 DATA No",
        "Process 6 DATE",
        "Process 6 TIME",
        "Process 6 Model Code",
        "Process 6 S/N",
        "Process 6 ID",
        "Process 6 NAME",
        "Process 6 Regular/Contractual",
        "Process 6 Vinyl",
        "Process 6 Vinyl Lot No",
        "Process 6 ST",
        "Process 6 Actual Time",
        "Process 6 NG Cause",
        "Process 6 Repaired Action"]

    dfVt1 = dfVt1[dfVt1["Process 1 Regular/Contractual"].str.contains("REG", na = False)]
    # dfVt1 = dfVt1[(dfVt1["Process 1 DATE"].isin([dateToday]))]
    dfVt1 = dfVt1[(dfVt1["Process 1 DATE"].isin([dateToRead]))]

    dfVt2 = dfVt2[dfVt2["Process 2 Regular/Contractual"].str.contains("REG", na = False)]
    # dfVt2 = dfVt2[(dfVt2["Process 2 DATE"].isin([dateToday]))]
    dfVt2 = dfVt2[(dfVt2["Process 2 DATE"].isin([dateToRead]))]

    dfVt3 = dfVt3[dfVt3["Process 3 Regular/Contractual"].str.contains("REG", na = False)]
    # dfVt3 = dfVt3[(dfVt3["Process 3 DATE"].isin([dateToday]))]
    dfVt3 = dfVt3[(dfVt3["Process 3 DATE"].isin([dateToRead]))]

    dfVt4 = dfVt4[dfVt4["Process 4 Regular/Contractual"].str.contains("REG", na = False)]
    # dfVt4 = dfVt4[(dfVt4["Process 4 DATE"].isin([dateToday]))]
    dfVt4 = dfVt4[(dfVt4["Process 4 DATE"].isin([dateToRead]))]

    dfVt5 = dfVt5[dfVt5["Process 5 Regular/Contractual"].str.contains("REG", na = False)]
    # dfVt5 = dfVt5[(dfVt5["Process 5 DATE"].isin([dateToday]))]
    dfVt5 = dfVt5[(dfVt5["Process 5 DATE"].isin([dateToRead]))]

    dfVt6 = dfVt6[dfVt6["Process 6 Regular/Contractual"].str.contains("REG", na = False)]
    # dfVt6 = dfVt6[(dfVt6["Process 6 DATE"].isin([dateToday]))]
    dfVt6 = dfVt6[(dfVt6["Process 6 DATE"].isin([dateToRead]))]

# %%
def CsvOrganize():
    global dfVt1
    global dfVt2
    global dfVt3
    global dfVt4
    global dfVt5
    global dfVt6

    global process1Row
    global process2Row
    global process3Row
    global process4Row
    global process5Row
    global process6Row

    global tempDfVt1
    global tempDfVt2
    global tempDfVt3
    global tempDfVt4
    global tempDfVt5
    global tempDfVt6

    global dfPi
    global tempdfPi
    global piRow

    global ngProcess
    
    global process1Status
    global process2Status
    global process3Status
    global process4Status
    global process5Status
    global process6Status
    global isRepairedWithNG
    global piStatus

    global canCompile

    global programRunning

    ngProcess = "-"

    process1Status = ""
    process2Status = ""
    process3Status = ""
    process4Status = ""
    process5Status = ""
    process6Status = ""
    isRepairedWithNG = False
    piStatus = ""

    isVt1Blank = False
    isVt2Blank = False
    isVt3Blank = False
    isVt4Blank = False
    isVt5Blank = False
    isVt6Blank = False

    # ReadPI In PiRow Value
    try:
        tempdfPi = dfPi.iloc[[piRow], :]
    except IndexError:
        pass

    if "INSPECTION ONLY" in tempdfPi["PROCESS S/N"].values:
        piStatus = "INSPECTION ONLY"
    else:

        try:
            #Checking If There's Value In tempDfVt1 To 6
            tempDfVt1 = dfVt1.iloc[[process1Row], :]
            tempDfVt2 = dfVt2.iloc[[process2Row], :]
            tempDfVt3 = dfVt3.iloc[[process3Row], :]
            tempDfVt4 = dfVt4.iloc[[process4Row], :]
            tempDfVt5 = dfVt5.iloc[[process5Row], :]
            tempDfVt6 = dfVt6.iloc[[process6Row], :]

            if tempDfVt1["Process 1 Repaired Action"].values[0] == "-" and tempDfVt2["Process 2 Repaired Action"].values[0] == "-" and tempDfVt3["Process 3 Repaired Action"].values[0] == "-" and tempDfVt4["Process 4 Repaired Action"].values[0] == "-" and tempDfVt5["Process 5 Repaired Action"].values[0] == "-" and tempDfVt6["Process 6 Repaired Action"].values[0] == "-":
                if tempDfVt1["Process 1 NG Cause"].values[0] == "-":
                    print("Process1 Good")
                    process1Status = "Good"
                    if tempDfVt2["Process 2 NG Cause"].values[0] == "-":
                        print("Process2 Good")
                        process2Status = "Good"
                        if tempDfVt3["Process 3 NG Cause"].values[0] == "-":
                            print("Process3 Good")
                            process3Status = "Good"
                            if tempDfVt4["Process 4 NG Cause"].values[0] == "-":
                                print("Process4 Good")
                                process4Status = "Good"
                                if tempDfVt5["Process 5 NG Cause"].values[0] == "-":
                                    print("Process5 Good")
                                    process5Status = "Good"
                                    if tempDfVt6["Process 6 NG Cause"].values[0] == "-":
                                        print("Process6 Good")
                                        process6Status = "Good"
                                    else:
                                        print("Process6 NG")
                                        process6Status = "NG"
                                elif tempDfVt5["Process 5 NG Cause"].values[0].replace(' ', '') == "NGPRESSURE":
                                    print("Process5 NG PRESSURE")
                                    process5Status = "NG PRESSURE"
                                else:
                                    print("Process5 NG")
                                    process5Status = "NG"
                            else:
                                print("Process4 NG")
                                process4Status = "NG"
                        else:
                            print("Process3 NG")
                            process3Status = "NG"
                    else:
                        print("Process2 NG")
                        process2Status = "NG"
                else:
                    print("Process1 NG")
                    process1Status = "NG"
            else:
                print("Repaired")
                
                if tempDfVt1["Process 1 Repaired Action"].values[0] != "-":
                    process1Status = "Repaired"
                    process2Status = ""
                    process3Status = ""
                    process4Status = ""
                    process5Status = ""
                    process6Status = ""
                if tempDfVt2["Process 2 Repaired Action"].values[0] != "-":
                    process1Status = ""
                    process2Status = "Repaired"
                    process3Status = ""
                    process4Status = ""
                    process5Status = ""
                    process6Status = ""
                if tempDfVt3["Process 3 Repaired Action"].values[0] != "-":
                    process1Status = ""
                    process2Status = ""
                    process3Status = "Repaired"
                    process4Status = ""
                    process5Status = ""
                    process6Status = ""
                if tempDfVt4["Process 4 Repaired Action"].values[0] != "-":
                    process1Status = ""
                    process2Status = ""
                    process3Status = ""
                    process4Status = "Repaired"
                    process5Status = ""
                    process6Status = ""
                if tempDfVt5["Process 5 Repaired Action"].values[0] != "-":
                    process1Status = ""
                    process2Status = ""
                    process3Status = ""
                    process4Status = ""
                    process5Status = "Repaired"
                    process6Status = ""
                if tempDfVt6["Process 6 Repaired Action"].values[0] != "-":
                    process1Status = ""
                    process2Status = ""
                    process3Status = ""
                    process4Status = ""
                    process5Status = ""
                    process6Status = "Repaired"

                #Checking Again For NG Process
                if tempDfVt1["Process 1 NG Cause"].values[0] != "-":
                    process1Status = "NG"
                    isRepairedWithNG = True
                elif tempDfVt2["Process 2 NG Cause"].values[0] != "-":
                    process2Status = "NG"
                    isRepairedWithNG = True
                elif tempDfVt3["Process 3 NG Cause"].values[0] != "-":
                    process3Status = "NG"
                    isRepairedWithNG = True
                elif tempDfVt4["Process 4 NG Cause"].values[0] != "-":
                    process4Status = "NG"
                    isRepairedWithNG = True
                elif tempDfVt5["Process 5 NG Cause"].values[0].replace(' ', '') == "NGPRESSURE":
                    print("Repaired With NG PRESSURE__________________________________________________________________________________________________________________")
                    process5Status = "NG PRESSURE"
                    isRepairedWithNG = True
                elif tempDfVt5["Process 5 NG Cause"].values[0] != "-":
                    process5Status = "NG"
                    isRepairedWithNG = True
                elif tempDfVt6["Process 6 NG Cause"].values[0] != "-":
                    process6Status = "NG"
                    isRepairedWithNG = True
            canCompile = True
        except:
            #Checking What tempDfVt Is Blank
            try:
                tempDfVt1 = dfVt1.iloc[[process1Row], :]
                isVt1Blank = False
            except:
                print("VT1 Blank")
                isVt1Blank = True
            try:
                tempDfVt2 = dfVt2.iloc[[process2Row], :]
                isVt2Blank = False
            except:
                print("VT2 Blank")
                isVt2Blank = True
            try:
                tempDfVt3 = dfVt3.iloc[[process3Row], :]
                isVt3Blank = False
            except:
                print("VT3 Blank")
                isVt3Blank = True
            try:
                tempDfVt4 = dfVt4.iloc[[process4Row], :]
                isVt4Blank = False
            except:
                print("VT4 Blank")
                isVt4Blank = True
            try:
                tempDfVt5 = dfVt5.iloc[[process5Row], :]
                isVt5Blank = False
            except:
                print("VT5 Blank")
                isVt5Blank = True
            try:
                tempDfVt6 = dfVt6.iloc[[process6Row], :]
                isVt6Blank = False
            except:
                print("VT6 Blank")
                isVt6Blank = True
            #No Data In Next Row
            if isVt1Blank == True and isVt2Blank == True and isVt3Blank == True and isVt4Blank == True and isVt5Blank == True and isVt6Blank == True:
                print("No More To Read")
                canCompile = False
            #Blank At Process2, Process3, Process4, Process5
            elif isVt1Blank == False and isVt2Blank == True and isVt3Blank == True and isVt4Blank == True and isVt5Blank == True and isVt6Blank == True and tempDfVt1["Process 1 Repaired Action"].values[0] == "-":
                if tempDfVt1["Process 1 NG Cause"].values[0] != "-":
                    print("Process 1 Proceed With NG")
                    process1Status = "NG"
                    canCompile = True
                else:
                    print("Pending In Process 1")
                    canCompile = False
            #Blank At Process3, Process4, Process 5
            elif isVt1Blank == False and isVt2Blank == False and isVt3Blank == True and isVt4Blank == True and isVt5Blank == True and isVt6Blank == True and tempDfVt2["Process 2 Repaired Action"].values[0] == "-":
                if tempDfVt2["Process 2 NG Cause"].values[0] != "-":
                    print("Process 2 Proceed With NG")
                    process1Status = "Good"
                    process2Status = "NG"
                    canCompile = True
                else:
                    print("Pending In Process 1 and Process 2")
                    canCompile = False
            #Blank At Process4, Process5
            elif isVt1Blank == False and isVt2Blank == False and isVt3Blank == False and isVt4Blank == True and isVt5Blank == True and isVt6Blank == True and tempDfVt3["Process 3 Repaired Action"].values[0] == "-":
                if tempDfVt3["Process 3 NG Cause"].values[0] != "-":
                    print("Process 3 Proceed With NG")
                    process1Status = "Good"
                    process2Status = "Good"
                    process3Status = "NG"
                    canCompile = True
                else:
                    print("Pending In Process 1 and Process 2 and Process 3")
                    canCompile = False
            #Blank At Process5
            elif isVt1Blank == False and isVt2Blank == False and isVt3Blank == False and isVt4Blank == False and isVt5Blank == True and isVt6Blank == True and tempDfVt4["Process 4 Repaired Action"].values[0] == "-":
                if tempDfVt4["Process 4 NG Cause"].values[0] != "-":
                    print("Process 4 Proceed With NG")
                    process1Status = "Good"
                    process2Status = "Good"
                    process3Status = "Good"
                    process4Status = "NG"
                    canCompile = True
                else:
                    print("Pending In Process 1 and Process 2 and Process 3 and Process 4")
                    canCompile = False
            #Blank At Process6       
            elif isVt1Blank == False and isVt2Blank == False and isVt3Blank == False and isVt4Blank == False and isVt5Blank == False and isVt6Blank == True and tempDfVt4["Process 4 Repaired Action"].values[0] == "-":
                if tempDfVt5["Process 5 NG Cause"].values[0] != "-":
                    print("Process 5 Proceed With NG")
                    process1Status = "Good"
                    process2Status = "Good"
                    process3Status = "Good"
                    process4Status = "Good"
                    process5Status = "NG"
                    canCompile = True
                else:
                    print("Pending In Process 1 and Process 2 and Process 3 and Process 4 and Process 5")
                    canCompile = False


            #Repair Process 1
            elif isVt1Blank == False and isVt2Blank == True and isVt3Blank == True and isVt4Blank == True and isVt5Blank == True and isVt6Blank == True and tempDfVt1["Process 1 Repaired Action"].values[0] != "-":
                    print("Pending Repair At Process 1")
                    canCompile = False
            elif isVt1Blank == False and isVt2Blank == False and isVt3Blank == True and isVt4Blank == True and isVt5Blank == True and isVt6Blank == True and tempDfVt1["Process 1 Repaired Action"].values[0] != "-":
                    print("Pending Repair At Process 1, Pending In Process 2")
                    canCompile = False
            elif isVt1Blank == False and isVt2Blank == False and isVt3Blank == False and isVt4Blank == True and isVt5Blank == True and isVt6Blank == True and tempDfVt1["Process 1 Repaired Action"].values[0] != "-":
                    print("Pending Repair At Process 1, Pending In Process 2, Pending In Process 3")
                    canCompile = False
            elif isVt1Blank == False and isVt2Blank == False and isVt3Blank == False and isVt4Blank == False and isVt5Blank == True and isVt6Blank == True and tempDfVt1["Process 1 Repaired Action"].values[0] != "-":
                    print("Pending Repair At Process 1, Pending In Process 2, Pending In Process 3, Pending In Process 4")
                    canCompile = False
            elif isVt1Blank == False and isVt2Blank == False and isVt3Blank == False and isVt4Blank == False and isVt5Blank == False and isVt6Blank == True and tempDfVt1["Process 1 Repaired Action"].values[0] != "-":
                    print("Pending Repair At Process 1, Pending In Process 2, Pending In Process 3, Pending In Process 4, Pending In Process 6")
                    canCompile = False
            #Repair Process 2
            elif isVt1Blank == True and isVt2Blank == False and isVt3Blank == True and isVt4Blank == True and isVt5Blank == True and isVt6Blank == True and tempDfVt2["Process 2 Repaired Action"].values[0] != "-":
                    print("Pending Repair At Process 2")
                    canCompile = False
            elif isVt1Blank == True and isVt2Blank == False and isVt3Blank == False and isVt4Blank == True and isVt5Blank == True and isVt6Blank == True and tempDfVt2["Process 2 Repaired Action"].values[0] != "-":
                    print("Pending Repair At Process 2, Process 3")
                    canCompile = False
            elif isVt1Blank == True and isVt2Blank == False and isVt3Blank == False and isVt4Blank == False and isVt5Blank == True and isVt6Blank == True and tempDfVt2["Process 2 Repaired Action"].values[0] != "-":
                    print("Pending Repair At Process 2, Process 3, Process 4")
                    canCompile = False
            elif isVt1Blank == True and isVt2Blank == False and isVt3Blank == False and isVt4Blank == False and isVt5Blank == False and isVt6Blank == True and tempDfVt2["Process 2 Repaired Action"].values[0] != "-":
                    print("Pending Repair At Process 2, Process 3, Process 4, Process 5")
                    canCompile = False
            elif isVt1Blank == True and isVt2Blank == False and isVt3Blank == False and isVt4Blank == False and isVt5Blank == False and isVt6Blank == False and tempDfVt2["Process 2 Repaired Action"].values[0] != "-":
                    process2Status = "Repaired"
                    canCompile = True
            #Repair Process 3
            elif isVt1Blank == True and isVt2Blank == True and isVt3Blank == False and isVt4Blank == True and isVt5Blank == True and isVt6Blank == True and tempDfVt3["Process 3 Repaired Action"].values[0] != "-":
                    print("Pending Repair At Process 3")
                    canCompile = False
            elif isVt1Blank == True and isVt2Blank == True and isVt3Blank == False and isVt4Blank == False and isVt5Blank == True and isVt6Blank == True and tempDfVt3["Process 3 Repaired Action"].values[0] != "-":
                    print("Pending Repair At Process 3, Process 4")
                    canCompile = False
            elif isVt1Blank == True and isVt2Blank == True and isVt3Blank == False and isVt4Blank == False and isVt5Blank == False and isVt6Blank == True and tempDfVt3["Process 3 Repaired Action"].values[0] != "-":
                    print("Pending Repair At Process 3, Process 4, Process 5")
                    canCompile = False
            elif isVt1Blank == True and isVt2Blank == True and isVt3Blank == False and isVt4Blank == False and isVt5Blank == False and isVt6Blank == False and tempDfVt3["Process 3 Repaired Action"].values[0] != "-":
                    process3Status = "Repaired"
                    canCompile = True
            #Repair Process 4
            elif isVt1Blank == True and isVt2Blank == True and isVt3Blank == True and isVt4Blank == False and isVt5Blank == True and isVt6Blank == True and tempDfVt4["Process 4 Repaired Action"].values[0] != "-":
                    print("Pending Repair At Process 4")
                    canCompile = False
            elif isVt1Blank == True and isVt2Blank == True and isVt3Blank == True and isVt4Blank == False and isVt5Blank == False and isVt6Blank == True and tempDfVt4["Process 4 Repaired Action"].values[0] != "-":
                    print("Pending Repair At Process 4, Process 5")
                    canCompile = False
            elif isVt1Blank == True and isVt2Blank == True and isVt3Blank == True and isVt4Blank == False and isVt5Blank == False and isVt6Blank == False and tempDfVt4["Process 4 Repaired Action"].values[0] != "-":
                    process4Status = "Repaired"
                    canCompile = True
            #Repair Process 5
            elif isVt1Blank == True and isVt2Blank == True and isVt3Blank == True and isVt4Blank == True and isVt5Blank == False and isVt6Blank == True and tempDfVt5["Process 5 Repaired Action"].values[0] != "-":
                    print("Pending Repair At Process 5")
                    canCompile = False
            elif isVt1Blank == True and isVt2Blank == True and isVt3Blank == True and isVt4Blank == True and isVt5Blank == False and isVt6Blank == False and tempDfVt5["Process 5 Repaired Action"].values[0] != "-":
                    process5Status = "Repaired"
                    canCompile = True
            #Repair Process 6
            elif isVt1Blank == True and isVt2Blank == True and isVt3Blank == True and isVt4Blank == True and isVt5Blank == True and isVt6Blank == False and tempDfVt6["Process 6 Repaired Action"].values[0] != "-":
                    process6Status = "Repaired"
                    canCompile = True
            else:
                canCompile = False

            if not canCompile:
                programRunning = False

# %%
def CompileCsv():
    global ngProcess
    global excelData
    global compiledFrame

    global process1Row
    global process2Row
    global process3Row
    global process4Row
    global process5Row
    global process6Row

    global process1Status
    global process2Status
    global process3Status
    global process4Status
    global process5Status
    global process6Status
    global isRepairedWithNG
    global piStatus

    global tempDfVt1
    global tempDfVt2
    global tempDfVt3
    global tempDfVt4
    global tempDfVt5
    global tempDfVt6

    global dfPi
    global tempdfPi
    global piRow

    # ReadPI In PiRow Value
    try:
        tempdfPi = dfPi.iloc[[piRow], :]
    except IndexError:
        pass

    #GETTING DATE TODAY
    DateAndTimeManager.GetDateToday()

    # GETTING EM2P INSPECTION DATA
    em2p = em2P()
    em2p.GettingData(tempDfVt1["Process 1 Em2p"].values, tempDfVt1["Process 1 Em2p Lot No"].values)

    #GETTING EM3P INSPECTION DATA
    em3p = em3P()
    em3p.GettingData(tempDfVt1["Process 1 Em3p"].values, tempDfVt1["Process 1 Em3p Lot No"].values)

    #GETTING FM INSPECTION DATA
    fm = fM()
    fm.GettingData(tempDfVt1["Process 1 Frame"].values, tempDfVt1["Process 1 Frame Lot No"].values)

    #GETTING DFB INSPECTION DATA
    dfb = dFB()
    dfb.ReadDfbSnap(tempDfVt2["Process 2 Df Blk Lot No"].values[0])
    dfb.GettingData(tempDfVt2["Process 2 Df Blk"].values[0])

    #GETTING TENSILE FOR DFB
    tensile = Tensile()
    tensile.GettingData(tempDfVt2["Process 2 Df Blk"].values[0], dfb.dfbLotNumber2[:-3])

    #GETTING RDB INSPECTION DATA
    rdb = rDB()
    rdb.ReadCheckSheet(tempDfVt2["Process 2 Rod Blk Lot No"].values[0], tempDfVt2["Process 2 Rod Blk"].values)
    rdb.GettingData(tempDfVt2["Process 2 Rod Blk"].values)

    #GETTING CSB INSPECTION DATA
    csb = cSB()
    csb.GettingData(tempDfVt3["Process 3 Casing Block"].values, tempDfVt3["Process 3 Casing Block Lot No"].values)

    excelData = {
        "DATE": tempdfPi["DATE"].values,
        "TIME": tempdfPi["TIME"].values,
        "MODEL CODE": tempdfPi["MODEL CODE"].values,
        "PROCESS S/N": tempdfPi["PROCESS S/N"].values,
        "S/N": tempdfPi["S/N"].values,
        "PASS/NG": tempdfPi["PASS/NG"].values,
        "VOLTAGE MAX (V)": tempdfPi["VOLTAGE MAX (V)"].values,
        "WATTAGE MAX (W)": tempdfPi["WATTAGE MAX (W)"].values,
        "CLOSED PRESSURE_MAX (kPa)": tempdfPi["CLOSED PRESSURE_MAX (kPa)"].values,
        "VOLTAGE Middle (V)": tempdfPi["VOLTAGE Middle (V)"].values,
        "WATTAGE Middle (W)": tempdfPi["WATTAGE Middle (W)"].values,
        "AMPERAGE Middle (A)": tempdfPi["AMPERAGE Middle (A)"].values,
        "CLOSED PRESSURE Middle (kPa)": tempdfPi["CLOSED PRESSURE Middle (kPa)"].values,
        "dB(A) 1": tempdfPi["dB(A) 1"].values,
        "dB(A) 2": tempdfPi["dB(A) 2"].values,
        "dB(A) 3": tempdfPi["dB(A) 3"].values,
        "VOLTAGE MIN (V)": tempdfPi["VOLTAGE MIN (V)"].values,
        "WATTAGE MIN (W)": tempdfPi["WATTAGE MIN (W)"].values,
        "CLOSED PRESSURE MIN (kPa)": tempdfPi["CLOSED PRESSURE MIN (kPa)"].values,

        "Process 1 Model Code": tempDfVt1["Process 1 Model Code"].values,
        "Process 1 S/N": tempDfVt1["Process 1 S/N"].values,
        "Process 1 ID": tempDfVt1["Process 1 ID"].values,
        "Process 1 NAME": tempDfVt1["Process 1 NAME"].values,
        "Process 1 Regular/Contractual": tempDfVt1["Process 1 Regular/Contractual"].values,
        "Process 1 Em2p": tempDfVt1["Process 1 Em2p"].values,
        "Process 1 Em2p Lot No": tempDfVt1["Process 1 Em2p Lot No"].values,
        "Process 1 Em2p Inspection 3 Average Data": em2p.totalAverage3,
        "Process 1 Em2p Inspection 4 Average Data": em2p.totalAverage4,
        "Process 1 Em2p Inspection 5 Average Data": em2p.totalAverage5,
        "Process 1 Em2p Inspection 10 Average Data": em2p.totalAverage10,
        "Process 1 Em2p Inspection 3 Minimum Data": em2p.totalMinimum3,
        "Process 1 Em2p Inspection 4 Minimum Data": em2p.totalMinimum4,
        "Process 1 Em2p Inspection 5 Minimum Data": em2p.totalMinimum5,
        "Process 1 Em2p Inspection 3 Maximum Data": em2p.totalMaximum3,
        "Process 1 Em2p Inspection 4 Maximum Data": em2p.totalMaximum4,
        "Process 1 Em2p Inspection 5 Maximum Data": em2p.totalMaximum5,
        "Process 1 Em3p": tempDfVt1["Process 1 Em3p"].values,
        "Process 1 Em3p Lot No": tempDfVt1["Process 1 Em3p Lot No"].values,
        "Process 1 Em3p Inspection 3 Average Data": em3p.totalAverage3,
        "Process 1 Em3p Inspection 4 Average Data": em3p.totalAverage4,
        "Process 1 Em3p Inspection 5 Average Data": em3p.totalAverage5,
        "Process 1 Em3p Inspection 10 Average Data": em3p.totalAverage10,
        "Process 1 Em3p Inspection 3 Minimum Data": em3p.totalMinimum3,
        "Process 1 Em3p Inspection 4 Minimum Data": em3p.totalMinimum4,
        "Process 1 Em3p Inspection 5 Minimum Data": em3p.totalMinimum5,
        "Process 1 Em3p Inspection 3 Maximum Data": em3p.totalMaximum3,
        "Process 1 Em3p Inspection 4 Maximum Data": em3p.totalMaximum4,
        "Process 1 Em3p Inspection 5 Maximum Data": em3p.totalMaximum5,
        "Process 1 Harness": tempDfVt1["Process 1 Harness"].values,
        "Process 1 Harness Lot No": tempDfVt1["Process 1 Harness Lot No"].values,
        "Process 1 Frame": tempDfVt1["Process 1 Frame"].values,
        "Process 1 Frame Lot No": tempDfVt1["Process 1 Frame Lot No"].values,
        "Process 1 Frame Inspection 1 Average Data": fm.totalAverage1, 
        "Process 1 Frame Inspection 2 Average Data": fm.totalAverage2, 
        "Process 1 Frame Inspection 3 Average Data": fm.totalAverage3, 
        "Process 1 Frame Inspection 4 Average Data": fm.totalAverage4, 
        "Process 1 Frame Inspection 5 Average Data": fm.totalAverage5, 
        "Process 1 Frame Inspection 6 Average Data": fm.totalAverage6, 
        "Process 1 Frame Inspection 7 Average Data": fm.totalAverage7, 
        "Process 1 Frame Inspection 1 Minimum Data": fm.totalMinimum1, 
        "Process 1 Frame Inspection 2 Minimum Data": fm.totalMinimum2, 
        "Process 1 Frame Inspection 3 Minimum Data": fm.totalMinimum3, 
        "Process 1 Frame Inspection 4 Minimum Data": fm.totalMinimum4, 
        "Process 1 Frame Inspection 5 Minimum Data": fm.totalMinimum5, 
        "Process 1 Frame Inspection 6 Minimum Data": fm.totalMinimum6, 
        "Process 1 Frame Inspection 7 Minimum Data": fm.totalMinimum7, 
        "Process 1 Frame Inspection 1 Maximum Data": fm.totalMaximum1, 
        "Process 1 Frame Inspection 2 Maximum Data": fm.totalMaximum2, 
        "Process 1 Frame Inspection 3 Maximum Data": fm.totalMaximum3, 
        "Process 1 Frame Inspection 4 Maximum Data": fm.totalMaximum4, 
        "Process 1 Frame Inspection 5 Maximum Data": fm.totalMaximum5, 
        "Process 1 Frame Inspection 6 Maximum Data": fm.totalMaximum6, 
        "Process 1 Frame Inspection 7 Maximum Data": fm.totalMaximum7, 
        "Process 1 Bushing": tempDfVt1["Process 1 Bushing"].values,
        "Process 1 Bushing Lot No": tempDfVt1["Process 1 Bushing Lot No"].values,
        "Process 1 ST": tempDfVt1["Process 1 ST"].values,
        "Process 1 Actual Time": tempDfVt1["Process 1 Actual Time"].values,
        "Process 1 NG Cause": tempDfVt1["Process 1 NG Cause"].values,
        "Process 1 Repaired Action": tempDfVt1["Process 1 Repaired Action"].values,

        "Process 2 Model Code": tempDfVt2["Process 2 Model Code"].values,
        "Process 2 S/N": tempDfVt2["Process 2 S/N"].values,
        "Process 2 ID": tempDfVt2["Process 2 ID"].values,
        "Process 2 NAME": tempDfVt2["Process 2 NAME"].values,
        "Process 2 Regular/Contractual": tempDfVt2["Process 2 Regular/Contractual"].values,
        "Process 2 M4x40 Screw": tempDfVt2["Process 2 M4x40 Screw"].values,
        "Process 2 M4x40 Screw Lot No": tempDfVt2["Process 2 M4x40 Screw Lot No"].values,
        "Process 2 Rod Blk": tempDfVt2["Process 2 Rod Blk"].values,
        "Process 2 Rod Blk Lot No": tempDfVt2["Process 2 Rod Blk Lot No"].values,
        "Process 2 Rod Blk Tesla 1 Average Data": rdb.rdbTeslaTotalAverage1,
        "Process 2 Rod Blk Tesla 2 Average Data": rdb.rdbTeslaTotalAverage2,
        "Process 2 Rod Blk Tesla 3 Average Data": rdb.rdbTeslaTotalAverage3,
        "Process 2 Rod Blk Tesla 4 Average Data": rdb.rdbTeslaTotalAverage4,
        "Process 2 Rod Blk Tesla 1 Minimum Data": rdb.rdbTeslaTotalMinimum1,
        "Process 2 Rod Blk Tesla 2 Minimum Data": rdb.rdbTeslaTotalMinimum2,
        "Process 2 Rod Blk Tesla 3 Minimum Data": rdb.rdbTeslaTotalMinimum3,
        "Process 2 Rod Blk Tesla 4 Minimum Data": rdb.rdbTeslaTotalMinimum4,
        "Process 2 Rod Blk Tesla 1 Maximum Data": rdb.rdbTeslaTotalMaximum1,
        "Process 2 Rod Blk Tesla 2 Maximum Data": rdb.rdbTeslaTotalMaximum2,
        "Process 2 Rod Blk Tesla 3 Maximum Data": rdb.rdbTeslaTotalMaximum3,
        "Process 2 Rod Blk Tesla 4 Maximum Data": rdb.rdbTeslaTotalMaximum4,
        "Process 2 Rod Blk Inspection 1 Average Data": rdb.rdbTotalAverage1,
        "Process 2 Rod Blk Inspection 2 Average Data": rdb.rdbTotalAverage2,
        "Process 2 Rod Blk Inspection 3 Average Data": rdb.rdbTotalAverage3,
        "Process 2 Rod Blk Inspection 4 Average Data": rdb.rdbTotalAverage4,
        "Process 2 Rod Blk Inspection 5 Average Data": rdb.rdbTotalAverage5,
        "Process 2 Rod Blk Inspection 6 Average Data": rdb.rdbTotalAverage6,
        "Process 2 Rod Blk Inspection 7 Average Data": rdb.rdbTotalAverage7,
        "Process 2 Rod Blk Inspection 8 Average Data": rdb.rdbTotalAverage8,
        "Process 2 Rod Blk Inspection 9 Average Data": rdb.rdbTotalAverage9,
        "Process 2 Rod Blk Inspection 1 Minimum Data": rdb.rdbTotalMinimum1,
        "Process 2 Rod Blk Inspection 2 Minimum Data": rdb.rdbTotalMinimum2,
        "Process 2 Rod Blk Inspection 3 Minimum Data": rdb.rdbTotalMinimum3,
        "Process 2 Rod Blk Inspection 4 Minimum Data": rdb.rdbTotalMinimum4,
        "Process 2 Rod Blk Inspection 5 Minimum Data": rdb.rdbTotalMinimum5,
        "Process 2 Rod Blk Inspection 6 Minimum Data": rdb.rdbTotalMinimum6,
        "Process 2 Rod Blk Inspection 7 Minimum Data": rdb.rdbTotalMinimum7,
        "Process 2 Rod Blk Inspection 8 Minimum Data": rdb.rdbTotalMinimum8,
        "Process 2 Rod Blk Inspection 9 Minimum Data": rdb.rdbTotalMinimum9,
        "Process 2 Rod Blk Inspection 1 Maximum Data": rdb.rdbTotalMaximum1,
        "Process 2 Rod Blk Inspection 2 Maximum Data": rdb.rdbTotalMaximum2,
        "Process 2 Rod Blk Inspection 3 Maximum Data": rdb.rdbTotalMaximum3,
        "Process 2 Rod Blk Inspection 4 Maximum Data": rdb.rdbTotalMaximum4,
        "Process 2 Rod Blk Inspection 5 Maximum Data": rdb.rdbTotalMaximum5,
        "Process 2 Rod Blk Inspection 6 Maximum Data": rdb.rdbTotalMaximum6,
        "Process 2 Rod Blk Inspection 7 Maximum Data": rdb.rdbTotalMaximum7,
        "Process 2 Rod Blk Inspection 8 Maximum Data": rdb.rdbTotalMaximum8,
        "Process 2 Rod Blk Inspection 9 Maximum Data": rdb.rdbTotalMaximum9,
        "Process 2 Df Blk": tempDfVt2["Process 2 Df Blk"].values,
        "Process 2 Df Blk Lot No": tempDfVt2["Process 2 Df Blk Lot No"].values,
        "Process 2 Df Blk Inspection 1 Average Data": dfb.totalAverage1,
        "Process 2 Df Blk Inspection 2 Average Data": dfb.totalAverage2,
        "Process 2 Df Blk Inspection 3 Average Data": dfb.totalAverage3,
        "Process 2 Df Blk Inspection 4 Average Data": dfb.totalAverage4,
        "Process 2 Df Blk Inspection 1 Minimum Data": dfb.totalMinimum1,
        "Process 2 Df Blk Inspection 2 Minimum Data": dfb.totalMinimum2,
        "Process 2 Df Blk Inspection 3 Minimum Data": dfb.totalMinimum3,
        "Process 2 Df Blk Inspection 4 Minimum Data": dfb.totalMinimum4,
        "Process 2 Df Blk Inspection 1 Maximum Data": dfb.totalMaximum1,
        "Process 2 Df Blk Inspection 2 Maximum Data": dfb.totalMaximum2,
        "Process 2 Df Blk Inspection 3 Maximum Data": dfb.totalMaximum3,
        "Process 2 Df Blk Inspection 4 Maximum Data": dfb.totalMaximum4,
        "Process 2 Df Blk Tensile Rate Of Change Average" : tensile.rateOfChangeTotalAverage,
        "Process 2 Df Blk Tensile Rate Of Change Minimum" : tensile.rateOfChangeTotalMinimum,
        "Process 2 Df Blk Tensile Rate Of Change Maximum" : tensile.rateOfChangeTotalMaximum,
        "Process 2 Df Blk Tensile Start Force Average" : tensile.startForceTotalAverage,
        "Process 2 Df Blk Tensile Start Force Minimum" : tensile.startForceTotalMinimum,
        "Process 2 Df Blk Tensile Start Force Maximum" : tensile.startForceTotalMaximum,
        "Process 2 Df Blk Tensile Terminating Force Average" : tensile.terminatingForceTotalAverage,
        "Process 2 Df Blk Tensile Terminating Force Minimum" : tensile.terminatingForceTotalMinimum,
        "Process 2 Df Blk Tensile Terminating Force Maximum" : tensile.terminatingForceTotalMaximum,
        "Process 2 Df Ring": tempDfVt2["Process 2 Df Ring"].values,
        "Process 2 Df Ring Lot No": tempDfVt2["Process 2 Df Ring Lot No"].values,
        "Process 2 Washer": tempDfVt2["Process 2 Washer"].values,
        "Process 2 Washer Lot No": tempDfVt2["Process 2 Washer Lot No"].values,
        "Process 2 Lock Nut": tempDfVt2["Process 2 Lock Nut"].values,
        "Process 2 Lock Nut Lot No": tempDfVt2["Process 2 Lock Nut Lot No"].values,
        "Process 2 ST": tempDfVt2["Process 2 ST"].values,
        "Process 2 Actual Time": tempDfVt2["Process 2 Actual Time"].values,
        "Process 2 NG Cause": tempDfVt2["Process 2 NG Cause"].values,
        "Process 2 Repaired Action": tempDfVt2["Process 2 Repaired Action"].values,

        "Process 3 Model Code": tempDfVt3["Process 3 Model Code"].values,
        "Process 3 S/N": tempDfVt3["Process 3 S/N"].values,
        "Process 3 ID": tempDfVt3["Process 3 ID"].values,
        "Process 3 NAME": tempDfVt3["Process 3 NAME"].values,
        "Process 3 Regular/Contractual": tempDfVt3["Process 3 Regular/Contractual"].values,
        "Process 3 Frame Gasket": tempDfVt3["Process 3 Frame Gasket"].values,
        "Process 3 Frame Gasket Lot No": tempDfVt3["Process 3 Frame Gasket Lot No"].values,
        "Process 3 Casing Block": tempDfVt3["Process 3 Casing Block"].values,
        "Process 3 Casing Block Lot No": tempDfVt3["Process 3 Casing Block Lot No"].values,
        "Process 3 Casing Block Inspection 1 Average Data": csb.totalAverage1,
        "Process 3 Casing Block Inspection 1 Minimum Data": csb.totalMinimum1,
        "Process 3 Casing Block Inspection 1 Maximum Data": csb.totalMaximum1,
        "Process 3 Casing Gasket": tempDfVt3["Process 3 Casing Gasket"].values,
        "Process 3 Casing Gasket Lot No": tempDfVt3["Process 3 Casing Gasket Lot No"].values,
        "Process 3 M4x16 Screw 1": tempDfVt3["Process 3 M4x16 Screw 1"].values,
        "Process 3 M4x16 Screw 1 Lot No": tempDfVt3["Process 3 M4x16 Screw 1 Lot No"].values,
        "Process 3 M4x16 Screw 2": tempDfVt3["Process 3 M4x16 Screw 2"].values,
        "Process 3 M4x16 Screw 2 Lot No": tempDfVt3["Process 3 M4x16 Screw 2 Lot No"].values,
        "Process 3 Ball Cushion": tempDfVt3["Process 3 Ball Cushion"].values,
        "Process 3 Ball Cushion Lot No": tempDfVt3["Process 3 Ball Cushion Lot No"].values,
        "Process 3 Frame Cover": tempDfVt3["Process 3 Frame Cover"].values,
        "Process 3 Frame Cover Lot No": tempDfVt3["Process 3 Frame Cover Lot No"].values,
        "Process 3 Partition Board": tempDfVt3["Process 3 Partition Board"].values,
        "Process 3 Partition Board Lot No": tempDfVt3["Process 3 Partition Board Lot No"].values,
        "Process 3 Built In Tube 1": tempDfVt3["Process 3 Built In Tube 1"].values,
        "Process 3 Built In Tube 1 Lot No": tempDfVt3["Process 3 Built In Tube 1 Lot No"].values,
        "Process 3 Built In Tube 2": tempDfVt3["Process 3 Built In Tube 2"].values,
        "Process 3 Built In Tube 2 Lot No": tempDfVt3["Process 3 Built In Tube 2 Lot No"].values,
        "Process 3 Head Cover": tempDfVt3["Process 3 Head Cover"].values,
        "Process 3 Head Cover Lot No": tempDfVt3["Process 3 Head Cover Lot No"].values,
        "Process 3 Casing Packing": tempDfVt3["Process 3 Casing Packing"].values,
        "Process 3 Casing Packing Lot No": tempDfVt3["Process 3 Casing Packing Lot No"].values,
        "Process 3 M4x12 Screw": tempDfVt3["Process 3 M4x12 Screw"].values,
        "Process 3 M4x12 Screw Lot No": tempDfVt3["Process 3 M4x12 Screw Lot No"].values,
        "Process 3 Csb L": tempDfVt3["Process 3 Csb L"].values,
        "Process 3 Csb L Lot No": tempDfVt3["Process 3 Csb L Lot No"].values,
        "Process 3 Csb R": tempDfVt3["Process 3 Csb R"].values,
        "Process 3 Csb R Lot No": tempDfVt3["Process 3 Csb R Lot No"].values,
        "Process 3 Head Packing": tempDfVt3["Process 3 Head Packing"].values,
        "Process 3 Head Packing Lot No": tempDfVt3["Process 3 Head Packing Lot No"].values,
        "Process 3 ST": tempDfVt3["Process 3 ST"].values,
        "Process 3 Actual Time": tempDfVt3["Process 3 Actual Time"].values,
        "Process 3 NG Cause": tempDfVt3["Process 3 NG Cause"].values,
        "Process 3 Repaired Action": tempDfVt3["Process 3 Repaired Action"].values,

        "Process 4 Model Code": tempDfVt4["Process 4 Model Code"].values,
        "Process 4 S/N": tempDfVt4["Process 4 S/N"].values,
        "Process 4 ID": tempDfVt4["Process 4 ID"].values,
        "Process 4 NAME": tempDfVt4["Process 4 NAME"].values,
        "Process 4 Regular/Contractual": tempDfVt4["Process 4 Regular/Contractual"].values,
        "Process 4 Tank": tempDfVt4["Process 4 Tank"].values,
        "Process 4 Tank Lot No": tempDfVt4["Process 4 Tank Lot No"].values,
        "Process 4 Upper Housing": tempDfVt4["Process 4 Upper Housing"].values,
        "Process 4 Upper Housing Lot No": tempDfVt4["Process 4 Upper Housing Lot No"].values,
        "Process 4 Cord Hook": tempDfVt4["Process 4 Cord Hook"].values,
        "Process 4 Cord Hook Lot No": tempDfVt4["Process 4 Cord Hook Lot No"].values,
        "Process 4 M4x16 Screw": tempDfVt4["Process 4 M4x16 Screw"].values,
        "Process 4 M4x16 Screw Lot No": tempDfVt4["Process 4 M4x16 Screw Lot No"].values,
        "Process 4 Tank Gasket": tempDfVt4["Process 4 Tank Gasket"].values,
        "Process 4 Tank Gasket Lot No": tempDfVt4["Process 4 Tank Gasket Lot No"].values,
        "Process 4 Tank Cover": tempDfVt4["Process 4 Tank Cover"].values,
        "Process 4 Tank Cover Lot No": tempDfVt4["Process 4 Tank Cover Lot No"].values,
        "Process 4 Housing Gasket": tempDfVt4["Process 4 Housing Gasket"].values,
        "Process 4 Housing Gasket Lot No": tempDfVt4["Process 4 Housing Gasket Lot No"].values,
        "Process 4 M4x40 Screw": tempDfVt4["Process 4 M4x40 Screw"].values,
        "Process 4 M4x40 Screw Lot No": tempDfVt4["Process 4 M4x40 Screw Lot No"].values,
        "Process 4 PartitionGasket": tempDfVt4["Process 4 PartitionGasket"].values,
        "Process 4 PartitionGasket Lot No": tempDfVt4["Process 4 PartitionGasket Lot No"].values,
        "Process 4 M4x12 Screw": tempDfVt4["Process 4 M4x12 Screw"].values,
        "Process 4 M4x12 Screw Lot No": tempDfVt4["Process 4 M4x12 Screw Lot No"].values,
        "Process 4 Muffler": tempDfVt4["Process 4 Muffler"].values,
        "Process 4 Muffler Lot No": tempDfVt4["Process 4 Muffler Lot No"].values,
        "Process 4 Muffler Gasket": tempDfVt4["Process 4 Muffler Gasket"].values,
        "Process 4 Muffler Gasket Lot No": tempDfVt4["Process 4 Muffler Gasket Lot No"].values,
        "Process 4 VCR": tempDfVt4["Process 4 VCR"].values,
        "Process 4 VCR Lot No": tempDfVt4["Process 4 VCR Lot No"].values,
        "Process 4 ST": tempDfVt4["Process 4 ST"].values,
        "Process 4 Actual Time": tempDfVt4["Process 4 Actual Time"].values,
        "Process 4 NG Cause": tempDfVt4["Process 4 NG Cause"].values,
        "Process 4 Repaired Action": tempDfVt4["Process 4 Repaired Action"].values,
        
        "Process 5 Model Code": tempDfVt5["Process 5 Model Code"].values,
        "Process 5 S/N": tempDfVt5["Process 5 S/N"].values,
        "Process 5 ID": tempDfVt5["Process 5 ID"].values,
        "Process 5 NAME": tempDfVt5["Process 5 NAME"].values,
        "Process 5 Regular/Contractual": tempDfVt5["Process 5 Regular/Contractual"].values,
        "Process 5 Rating Label": tempDfVt5["Process 5 Rating Label"].values,
        "Process 5 Rating Label Lot No": tempDfVt5["Process 5 Rating Label Lot No"].values,
        "Process 5 ST": tempDfVt5["Process 5 ST"].values,
        "Process 5 Actual Time": tempDfVt5["Process 5 Actual Time"].values,
        "Process 5 NG Cause": tempDfVt5["Process 5 NG Cause"].values,
        "Process 5 Repaired Action": tempDfVt5["Process 5 Repaired Action"].values,
        
        "Process 6 Model Code": tempDfVt6["Process 6 Model Code"].values,
        "Process 6 S/N": tempDfVt6["Process 6 S/N"].values,
        "Process 6 ID": tempDfVt6["Process 6 ID"].values,
        "Process 6 NAME": tempDfVt6["Process 6 NAME"].values,
        "Process 6 Regular/Contractual": tempDfVt6["Process 6 Regular/Contractual"].values,
        "Process 6 Vinyl": tempDfVt6["Process 6 Vinyl"].values,
        "Process 6 Vinyl Lot No": tempDfVt6["Process 6 Vinyl Lot No"].values,
        "Process 6 ST": tempDfVt6["Process 6 ST"].values,
        "Process 6 Actual Time": tempDfVt6["Process 6 Actual Time"].values,
        "Process 6 NG Cause": tempDfVt6["Process 6 NG Cause"].values,
        "Process 6 Repaired Action": tempDfVt6["Process 6 Repaired Action"].values
    }
    excelData = pd.DataFrame(excelData)
    if piStatus == "INSPECTION ONLY":
        piRow += 1

        excelData["Process 1 Model Code"] = piStatus
        excelData["Process 1 S/N"] = piStatus
        excelData["Process 1 ID"] = piStatus
        excelData["Process 1 NAME"] = piStatus
        excelData["Process 1 Regular/Contractual"] = piStatus
        excelData["Process 1 Em2p"] = piStatus
        excelData["Process 1 Em2p Lot No"] = piStatus
        excelData["Process 1 Em3p"] = piStatus
        excelData["Process 1 Em3p Lot No"] = piStatus
        excelData["Process 1 Harness"] = piStatus
        excelData["Process 1 Harness Lot No"] = piStatus
        excelData["Process 1 Frame"] = piStatus
        excelData["Process 1 Frame Lot No"] = piStatus 
        excelData["Process 1 Bushing"] = piStatus
        excelData["Process 1 Bushing Lot No"] = piStatus
        excelData["Process 1 ST"] = piStatus
        excelData["Process 1 Actual Time"] = piStatus
        excelData["Process 1 NG Cause"] = piStatus
        excelData["Process 1 Repaired Action"] = piStatus 

        excelData["Process 2 Model Code"] = piStatus
        excelData["Process 2 S/N"] = piStatus
        excelData["Process 2 ID"] = piStatus
        excelData["Process 2 NAME"] = piStatus
        excelData["Process 2 Regular/Contractual"] = piStatus
        excelData["Process 2 M4x40 Screw"] = piStatus
        excelData["Process 2 M4x40 Screw Lot No"] = piStatus
        excelData["Process 2 Rod Blk"] = piStatus
        excelData["Process 2 Rod Blk Lot No"] = piStatus
        excelData["Process 2 Df Blk"] = piStatus
        excelData["Process 2 Df Blk Lot No"] = piStatus
        excelData["Process 2 Df Ring"] = piStatus
        excelData["Process 2 Df Ring Lot No"] = piStatus
        excelData["Process 2 Washer"] = piStatus
        excelData["Process 2 Washer Lot No"] = piStatus
        excelData["Process 2 Lock Nut"] = piStatus
        excelData["Process 2 Lock Nut Lot No"] = piStatus
        excelData["Process 2 ST"] = piStatus
        excelData["Process 2 Actual Time"] = piStatus
        excelData["Process 2 NG Cause"] = piStatus
        excelData["Process 2 Repaired Action"] = piStatus

        excelData["Process 3 Model Code"] = piStatus
        excelData["Process 3 S/N"] = piStatus
        excelData["Process 3 ID"] = piStatus
        excelData["Process 3 NAME"] = piStatus
        excelData["Process 3 Regular/Contractual"] = piStatus
        excelData["Process 3 Frame Gasket"] = piStatus
        excelData["Process 3 Frame Gasket Lot No"] = piStatus
        excelData["Process 3 Casing Block"] = piStatus
        excelData["Process 3 Casing Block Lot No"] = piStatus
        excelData["Process 3 Casing Block Inspection 1 Average Data"] = piStatus
        excelData["Process 3 Casing Block Inspection 1 Minimum Data"] = piStatus
        excelData["Process 3 Casing Block Inspection 1 Maximum Data"] = piStatus
        excelData["Process 3 Casing Gasket"] = piStatus
        excelData["Process 3 Casing Gasket Lot No"] = piStatus
        excelData["Process 3 M4x16 Screw 1"] = piStatus
        excelData["Process 3 M4x16 Screw 1 Lot No"] = piStatus
        excelData["Process 3 M4x16 Screw 2"] = piStatus
        excelData["Process 3 M4x16 Screw 2 Lot No"] = piStatus
        excelData["Process 3 Ball Cushion"] = piStatus
        excelData["Process 3 Ball Cushion Lot No"] = piStatus
        excelData["Process 3 Frame Cover"] = piStatus
        excelData["Process 3 Frame Cover Lot No"] = piStatus
        excelData["Process 3 Partition Board"] = piStatus
        excelData["Process 3 Partition Board Lot No"] = piStatus
        excelData["Process 3 Built In Tube 1"] = piStatus
        excelData["Process 3 Built In Tube 1 Lot No"] = piStatus
        excelData["Process 3 Built In Tube 2"] = piStatus
        excelData["Process 3 Built In Tube 2 Lot No"] = piStatus
        excelData["Process 3 Head Cover"] = piStatus
        excelData["Process 3 Head Cover Lot No"] = piStatus
        excelData["Process 3 Casing Packing"] = piStatus
        excelData["Process 3 Casing Packing Lot No"] = piStatus
        excelData["Process 3 M4x12 Screw"] = piStatus
        excelData["Process 3 M4x12 Screw Lot No"] = piStatus
        excelData["Process 3 Csb L"] = piStatus
        excelData["Process 3 Csb L Lot No"] = piStatus
        excelData["Process 3 Csb R"] = piStatus
        excelData["Process 3 Csb R Lot No"] = piStatus
        excelData["Process 3 Head Packing"] = piStatus
        excelData["Process 3 Head Packing Lot No"] = piStatus
        excelData["Process 3 ST"] = piStatus
        excelData["Process 3 Actual Time"] = piStatus
        excelData["Process 3 NG Cause"] = piStatus
        excelData["Process 3 Repaired Action"] = piStatus

        excelData["Process 4 Model Code"] = piStatus
        excelData["Process 4 S/N"] = piStatus
        excelData["Process 4 ID"] = piStatus
        excelData["Process 4 NAME"] = piStatus
        excelData["Process 4 Regular/Contractual"] = piStatus
        excelData["Process 4 Tank"] = piStatus
        excelData["Process 4 Tank Lot No"] = piStatus
        excelData["Process 4 Upper Housing"] = piStatus
        excelData["Process 4 Upper Housing Lot No"] = piStatus
        excelData["Process 4 Cord Hook" ] = piStatus
        excelData["Process 4 Cord Hook Lot No"] = piStatus
        excelData["Process 4 M4x16 Screw"] = piStatus
        excelData["Process 4 M4x16 Screw Lot No"] = piStatus
        excelData["Process 4 Tank Gasket"] = piStatus
        excelData["Process 4 Tank Gasket Lot No"] = piStatus
        excelData["Process 4 Tank Cover"] = piStatus
        excelData["Process 4 Tank Cover Lot No"] = piStatus
        excelData["Process 4 Housing Gasket"] = piStatus
        excelData["Process 4 Housing Gasket Lot No"] = piStatus
        excelData["Process 4 M4x40 Screw"] = piStatus
        excelData["Process 4 M4x40 Screw Lot No"] = piStatus
        excelData["Process 4 PartitionGasket"] = piStatus
        excelData["Process 4 PartitionGasket Lot No"] = piStatus
        excelData["Process 4 M4x12 Screw"] = piStatus
        excelData["Process 4 M4x12 Screw Lot No"] = piStatus
        excelData["Process 4 Muffler"] = piStatus
        excelData["Process 4 Muffler Lot No"] = piStatus
        excelData["Process 4 Muffler Gasket"] = piStatus
        excelData["Process 4 Muffler Gasket Lot No"] = piStatus
        excelData["Process 4 VCR"] = piStatus
        excelData["Process 4 VCR Lot No"] = piStatus
        excelData["Process 4 ST"] = piStatus
        excelData["Process 4 Actual Time"] = piStatus
        excelData["Process 4 NG Cause"] = piStatus
        excelData["Process 4 Repaired Action"] = piStatus
        
        excelData["Process 5 Model Code"] = piStatus
        excelData["Process 5 S/N"] = piStatus
        excelData["Process 5 ID"] = piStatus
        excelData["Process 5 NAME"] = piStatus
        excelData["Process 5 Regular/Contractual"] = piStatus
        excelData["Process 5 Rating Label"] = piStatus
        excelData["Process 5 Rating Label Lot No"] = piStatus
        excelData["Process 5 ST"] = piStatus
        excelData["Process 5 Actual Time"] = piStatus
        excelData["Process 5 NG Cause"] = piStatus
        excelData["Process 5 Repaired Action"] = piStatus
        
        excelData["Process 6 Model Code"] = piStatus
        excelData["Process 6 S/N"] = piStatus
        excelData["Process 6 ID"] = piStatus
        excelData["Process 6 NAME"] = piStatus
        excelData["Process 6 Regular/Contractual"] = piStatus
        excelData["Process 6 Vinyl"] = piStatus
        excelData["Process 6 Vinyl Lot No"] = piStatus
        excelData["Process 6 ST"] = piStatus
        excelData["Process 6 Actual Time"] = piStatus
        excelData["Process 6 NG Cause"] = piStatus
        excelData["Process 6 Repaired Action"] = piStatus
    else:
        if process1Status == "Good":
            process1Row += 1
        if process2Status == "Good":
            process2Row += 1
        if process3Status == "Good":
            process3Row += 1
        if process4Status == "Good":
            process4Row += 1
        if process5Status == "Good":
            process5Row += 1
            piRow += 1
        if process6Status == "Good":
            process6Row += 1
        
        if isRepairedWithNG:
            if process1Status == "Repaired":
                if process2Status == "NG":
                    ngProcess = "NG AT PROCESS2"
                    process1Row += 1
                    process2Row += 1

                    excelData["Process 3 Model Code"] = ngProcess
                    excelData["Process 3 S/N"] = ngProcess
                    excelData["Process 3 ID"] = ngProcess
                    excelData["Process 3 NAME"] = ngProcess
                    excelData["Process 3 Regular/Contractual"] = ngProcess
                    excelData["Process 3 Frame Gasket"] = ngProcess
                    excelData["Process 3 Frame Gasket Lot No"] = ngProcess
                    excelData["Process 3 Casing Block"] = ngProcess
                    excelData["Process 3 Casing Block Lot No"] = ngProcess
                    excelData["Process 3 Casing Block Inspection 1 Average Data"] = ngProcess
                    excelData["Process 3 Casing Block Inspection 1 Minimum Data"] = ngProcess
                    excelData["Process 3 Casing Block Inspection 1 Maximum Data"] = ngProcess
                    excelData["Process 3 Casing Gasket"] = ngProcess
                    excelData["Process 3 Casing Gasket Lot No"] = ngProcess
                    excelData["Process 3 M4x16 Screw 1"] = ngProcess
                    excelData["Process 3 M4x16 Screw 1 Lot No"] = ngProcess
                    excelData["Process 3 M4x16 Screw 2"] = ngProcess
                    excelData["Process 3 M4x16 Screw 2 Lot No"] = ngProcess
                    excelData["Process 3 Ball Cushion"] = ngProcess
                    excelData["Process 3 Ball Cushion Lot No"] = ngProcess
                    excelData["Process 3 Frame Cover"] = ngProcess
                    excelData["Process 3 Frame Cover Lot No"] = ngProcess
                    excelData["Process 3 Partition Board"] = ngProcess
                    excelData["Process 3 Partition Board Lot No"] = ngProcess
                    excelData["Process 3 Built In Tube 1"] = ngProcess
                    excelData["Process 3 Built In Tube 1 Lot No"] = ngProcess
                    excelData["Process 3 Built In Tube 2"] = ngProcess
                    excelData["Process 3 Built In Tube 2 Lot No"] = ngProcess
                    excelData["Process 3 Head Cover"] = ngProcess
                    excelData["Process 3 Head Cover Lot No"] = ngProcess
                    excelData["Process 3 Casing Packing"] = ngProcess
                    excelData["Process 3 Casing Packing Lot No"] = ngProcess
                    excelData["Process 3 M4x12 Screw"] = ngProcess
                    excelData["Process 3 M4x12 Screw Lot No"] = ngProcess
                    excelData["Process 3 Csb L"] = ngProcess
                    excelData["Process 3 Csb L Lot No"] = ngProcess
                    excelData["Process 3 Csb R"] = ngProcess
                    excelData["Process 3 Csb R Lot No"] = ngProcess
                    excelData["Process 3 Head Packing"] = ngProcess
                    excelData["Process 3 Head Packing Lot No"] = ngProcess
                    excelData["Process 3 ST"] = ngProcess
                    excelData["Process 3 Actual Time"] = ngProcess
                    excelData["Process 3 NG Cause"] = ngProcess
                    excelData["Process 3 Repaired Action"] = ngProcess

                    excelData["Process 4 Model Code"] = ngProcess
                    excelData["Process 4 S/N"] = ngProcess
                    excelData["Process 4 ID"] = ngProcess
                    excelData["Process 4 NAME"] = ngProcess
                    excelData["Process 4 Regular/Contractual"] = ngProcess
                    excelData["Process 4 Tank"] = ngProcess
                    excelData["Process 4 Tank Lot No"] = ngProcess
                    excelData["Process 4 Upper Housing"] = ngProcess
                    excelData["Process 4 Upper Housing Lot No"] = ngProcess
                    excelData["Process 4 Cord Hook" ] = ngProcess
                    excelData["Process 4 Cord Hook Lot No"] = ngProcess
                    excelData["Process 4 M4x16 Screw"] = ngProcess
                    excelData["Process 4 M4x16 Screw Lot No"] = ngProcess
                    excelData["Process 4 Tank Gasket"] = ngProcess
                    excelData["Process 4 Tank Gasket Lot No"] = ngProcess
                    excelData["Process 4 Tank Cover"] = ngProcess
                    excelData["Process 4 Tank Cover Lot No"] = ngProcess
                    excelData["Process 4 Housing Gasket"] = ngProcess
                    excelData["Process 4 Housing Gasket Lot No"] = ngProcess
                    excelData["Process 4 M4x40 Screw"] = ngProcess
                    excelData["Process 4 M4x40 Screw Lot No"] = ngProcess
                    excelData["Process 4 PartitionGasket"] = ngProcess
                    excelData["Process 4 PartitionGasket Lot No"] = ngProcess
                    excelData["Process 4 M4x12 Screw"] = ngProcess
                    excelData["Process 4 M4x12 Screw Lot No"] = ngProcess
                    excelData["Process 4 Muffler"] = ngProcess
                    excelData["Process 4 Muffler Lot No"] = ngProcess
                    excelData["Process 4 Muffler Gasket"] = ngProcess
                    excelData["Process 4 Muffler Gasket Lot No"] = ngProcess
                    excelData["Process 4 VCR"] = ngProcess
                    excelData["Process 4 VCR Lot No"] = ngProcess
                    excelData["Process 4 ST"] = ngProcess
                    excelData["Process 4 Actual Time"] = ngProcess
                    excelData["Process 4 NG Cause"] = ngProcess
                    excelData["Process 4 Repaired Action"] = ngProcess
                    
                    excelData["Process 5 Model Code"] = ngProcess
                    excelData["Process 5 S/N"] = ngProcess
                    excelData["Process 5 ID"] = ngProcess
                    excelData["Process 5 NAME"] = ngProcess
                    excelData["Process 5 Regular/Contractual"] = ngProcess
                    excelData["Process 5 Rating Label"] = ngProcess
                    excelData["Process 5 Rating Label Lot No"] = ngProcess
                    excelData["Process 5 ST"] = ngProcess
                    excelData["Process 5 Actual Time"] = ngProcess
                    excelData["Process 5 NG Cause"] = ngProcess
                    excelData["Process 5 Repaired Action"] = ngProcess
                    
                    excelData["Process 6 Model Code"] = ngProcess
                    excelData["Process 6 S/N"] = ngProcess
                    excelData["Process 6 ID"] = ngProcess
                    excelData["Process 6 NAME"] = ngProcess
                    excelData["Process 6 Regular/Contractual"] = ngProcess
                    excelData["Process 6 Vinyl"] = ngProcess
                    excelData["Process 6 Vinyl Lot No"] = ngProcess
                    excelData["Process 6 ST"] = ngProcess
                    excelData["Process 6 Actual Time"] = ngProcess
                    excelData["Process 6 NG Cause"] = ngProcess
                    excelData["Process 6 Repaired Action"] = ngProcess

                elif process3Status == "NG":
                    ngProcess = "NG AT PROCESS3"
                    process1Row += 1
                    process2Row += 1
                    process3Row += 1

                    excelData["Process 4 Model Code"] = ngProcess
                    excelData["Process 4 S/N"] = ngProcess
                    excelData["Process 4 ID"] = ngProcess
                    excelData["Process 4 NAME"] = ngProcess
                    excelData["Process 4 Regular/Contractual"] = ngProcess
                    excelData["Process 4 Tank"] = ngProcess
                    excelData["Process 4 Tank Lot No"] = ngProcess
                    excelData["Process 4 Upper Housing"] = ngProcess
                    excelData["Process 4 Upper Housing Lot No"] = ngProcess
                    excelData["Process 4 Cord Hook" ] = ngProcess
                    excelData["Process 4 Cord Hook Lot No"] = ngProcess
                    excelData["Process 4 M4x16 Screw"] = ngProcess
                    excelData["Process 4 M4x16 Screw Lot No"] = ngProcess
                    excelData["Process 4 Tank Gasket"] = ngProcess
                    excelData["Process 4 Tank Gasket Lot No"] = ngProcess
                    excelData["Process 4 Tank Cover"] = ngProcess
                    excelData["Process 4 Tank Cover Lot No"] = ngProcess
                    excelData["Process 4 Housing Gasket"] = ngProcess
                    excelData["Process 4 Housing Gasket Lot No"] = ngProcess
                    excelData["Process 4 M4x40 Screw"] = ngProcess
                    excelData["Process 4 M4x40 Screw Lot No"] = ngProcess
                    excelData["Process 4 PartitionGasket"] = ngProcess
                    excelData["Process 4 PartitionGasket Lot No"] = ngProcess
                    excelData["Process 4 M4x12 Screw"] = ngProcess
                    excelData["Process 4 M4x12 Screw Lot No"] = ngProcess
                    excelData["Process 4 Muffler"] = ngProcess
                    excelData["Process 4 Muffler Lot No"] = ngProcess
                    excelData["Process 4 Muffler Gasket"] = ngProcess
                    excelData["Process 4 Muffler Gasket Lot No"] = ngProcess
                    excelData["Process 4 VCR"] = ngProcess
                    excelData["Process 4 VCR Lot No"] = ngProcess
                    excelData["Process 4 ST"] = ngProcess
                    excelData["Process 4 Actual Time"] = ngProcess
                    excelData["Process 4 NG Cause"] = ngProcess
                    excelData["Process 4 Repaired Action"] = ngProcess
                    
                    excelData["Process 5 Model Code"] = ngProcess
                    excelData["Process 5 S/N"] = ngProcess
                    excelData["Process 5 ID"] = ngProcess
                    excelData["Process 5 NAME"] = ngProcess
                    excelData["Process 5 Regular/Contractual"] = ngProcess
                    excelData["Process 5 Rating Label"] = ngProcess
                    excelData["Process 5 Rating Label Lot No"] = ngProcess
                    excelData["Process 5 ST"] = ngProcess
                    excelData["Process 5 Actual Time"] = ngProcess
                    excelData["Process 5 NG Cause"] = ngProcess
                    excelData["Process 5 Repaired Action"] = ngProcess
                    
                    excelData["Process 6 Model Code"] = ngProcess
                    excelData["Process 6 S/N"] = ngProcess
                    excelData["Process 6 ID"] = ngProcess
                    excelData["Process 6 NAME"] = ngProcess
                    excelData["Process 6 Regular/Contractual"] = ngProcess
                    excelData["Process 6 Vinyl"] = ngProcess
                    excelData["Process 6 Vinyl Lot No"] = ngProcess
                    excelData["Process 6 ST"] = ngProcess
                    excelData["Process 6 Actual Time"] = ngProcess
                    excelData["Process 6 NG Cause"] = ngProcess
                    excelData["Process 6 Repaired Action"] = ngProcess

                elif process4Status == "NG":
                    ngProcess = "NG AT PROCESS4"
                    process1Row += 1
                    process2Row += 1
                    process3Row += 1
                    process4Row += 1

                    excelData["Process 5 Model Code"] = ngProcess
                    excelData["Process 5 S/N"] = ngProcess
                    excelData["Process 5 ID"] = ngProcess
                    excelData["Process 5 NAME"] = ngProcess
                    excelData["Process 5 Regular/Contractual"] = ngProcess
                    excelData["Process 5 Rating Label"] = ngProcess
                    excelData["Process 5 Rating Label Lot No"] = ngProcess
                    excelData["Process 5 ST"] = ngProcess
                    excelData["Process 5 Actual Time"] = ngProcess
                    excelData["Process 5 NG Cause"] = ngProcess
                    excelData["Process 5 Repaired Action"] = ngProcess
                    
                    excelData["Process 6 Model Code"] = ngProcess
                    excelData["Process 6 S/N"] = ngProcess
                    excelData["Process 6 ID"] = ngProcess
                    excelData["Process 6 NAME"] = ngProcess
                    excelData["Process 6 Regular/Contractual"] = ngProcess
                    excelData["Process 6 Vinyl"] = ngProcess
                    excelData["Process 6 Vinyl Lot No"] = ngProcess
                    excelData["Process 6 ST"] = ngProcess
                    excelData["Process 6 Actual Time"] = ngProcess
                    excelData["Process 6 NG Cause"] = ngProcess
                    excelData["Process 6 Repaired Action"] = ngProcess

                elif process5Status == "NG PRESSURE":
                    ngProcess = "NG PRESSURE AT PROCESS5"
                    process1Row += 1
                    process2Row += 1
                    process3Row += 1
                    process4Row += 1
                    process5Row += 1
                    piRow += 1

                    excelData["DATE"] = ngProcess
                    excelData["TIME"] = ngProcess
                    excelData["MODEL CODE"] = ngProcess
                    # excelData["PROCESS S/N"] = ngProcess
                    excelData["S/N"] = ngProcess
                    excelData["PASS/NG"] = ngProcess
                    excelData["VOLTAGE MAX (V)"] = ngProcess
                    excelData["WATTAGE MAX (W)"] = ngProcess
                    excelData["CLOSED PRESSURE_MAX (kPa)"] = ngProcess
                    excelData["VOLTAGE Middle (V)"] = ngProcess
                    excelData["WATTAGE Middle (W)"] = ngProcess
                    excelData["AMPERAGE Middle (A)"] = ngProcess
                    excelData["CLOSED PRESSURE Middle (kPa)"] = ngProcess
                    excelData["dB(A) 1"] = ngProcess
                    excelData["dB(A) 2"] = ngProcess
                    excelData["dB(A) 3"] = ngProcess
                    excelData["VOLTAGE MIN (V)"] = ngProcess
                    excelData["WATTAGE MIN (W)"] = ngProcess
                    excelData["CLOSED PRESSURE MIN (kPa)"] = ngProcess

                    excelData["Process 6 Model Code"] = ngProcess
                    excelData["Process 6 S/N"] = ngProcess
                    excelData["Process 6 ID"] = ngProcess
                    excelData["Process 6 NAME"] = ngProcess
                    excelData["Process 6 Regular/Contractual"] = ngProcess
                    excelData["Process 6 Vinyl"] = ngProcess
                    excelData["Process 6 Vinyl Lot No"] = ngProcess
                    excelData["Process 6 ST"] = ngProcess
                    excelData["Process 6 Actual Time"] = ngProcess
                    excelData["Process 6 NG Cause"] = ngProcess
                    excelData["Process 6 Repaired Action"] = ngProcess
                elif process5Status == "NG":
                    ngProcess = "NG AT PROCESS5"
                    process1Row += 1
                    process2Row += 1
                    process3Row += 1
                    process4Row += 1
                    process5Row += 1
                    piRow += 1

                    excelData["Process 6 Model Code"] = ngProcess
                    excelData["Process 6 S/N"] = ngProcess
                    excelData["Process 6 ID"] = ngProcess
                    excelData["Process 6 NAME"] = ngProcess
                    excelData["Process 6 Regular/Contractual"] = ngProcess
                    excelData["Process 6 Vinyl"] = ngProcess
                    excelData["Process 6 Vinyl Lot No"] = ngProcess
                    excelData["Process 6 ST"] = ngProcess
                    excelData["Process 6 Actual Time"] = ngProcess
                    excelData["Process 6 NG Cause"] = ngProcess
                    excelData["Process 6 Repaired Action"] = ngProcess

                elif process6Status == "NG":
                    ngProcess = "NG AT PROCESS6"
                    process1Row += 1
                    process2Row += 1
                    process3Row += 1
                    process4Row += 1
                    process5Row += 1
                    process6Row += 1
                    piRow += 1

            elif process2Status == "Repaired":
                if process3Status == "NG":
                    repairedProcess = "REPAIRED AT PROCESS2"
                    ngProcess = "NG AT PROCESS3"
                    process2Row += 1
                    process3Row += 1

                    excelData["Process 1 Model Code"] = repairedProcess
                    excelData["Process 1 S/N"] = repairedProcess
                    excelData["Process 1 ID"] = repairedProcess
                    excelData["Process 1 NAME"] = repairedProcess
                    excelData["Process 1 Regular/Contractual"] = repairedProcess
                    excelData["Process 1 Em2p"] = repairedProcess
                    excelData["Process 1 Em2p Lot No"] = repairedProcess
                    excelData["Process 1 Em3p"] = repairedProcess
                    excelData["Process 1 Em3p Lot No"] = repairedProcess
                    excelData["Process 1 Harness"] = repairedProcess
                    excelData["Process 1 Harness Lot No"] = repairedProcess
                    excelData["Process 1 Frame"] = repairedProcess
                    excelData["Process 1 Frame Lot No"] = repairedProcess 
                    excelData["Process 1 Bushing"] = repairedProcess
                    excelData["Process 1 Bushing Lot No"] = repairedProcess
                    excelData["Process 1 ST"] = repairedProcess
                    excelData["Process 1 Actual Time"] = repairedProcess
                    excelData["Process 1 NG Cause"] = repairedProcess
                    excelData["Process 1 Repaired Action"] = repairedProcess

                    excelData["Process 4 Model Code"] = ngProcess
                    excelData["Process 4 S/N"] = ngProcess
                    excelData["Process 4 ID"] = ngProcess
                    excelData["Process 4 NAME"] = ngProcess
                    excelData["Process 4 Regular/Contractual"] = ngProcess
                    excelData["Process 4 Tank"] = ngProcess
                    excelData["Process 4 Tank Lot No"] = ngProcess
                    excelData["Process 4 Upper Housing"] = ngProcess
                    excelData["Process 4 Upper Housing Lot No"] = ngProcess
                    excelData["Process 4 Cord Hook" ] = ngProcess
                    excelData["Process 4 Cord Hook Lot No"] = ngProcess
                    excelData["Process 4 M4x16 Screw"] = ngProcess
                    excelData["Process 4 M4x16 Screw Lot No"] = ngProcess
                    excelData["Process 4 Tank Gasket"] = ngProcess
                    excelData["Process 4 Tank Gasket Lot No"] = ngProcess
                    excelData["Process 4 Tank Cover"] = ngProcess
                    excelData["Process 4 Tank Cover Lot No"] = ngProcess
                    excelData["Process 4 Housing Gasket"] = ngProcess
                    excelData["Process 4 Housing Gasket Lot No"] = ngProcess
                    excelData["Process 4 M4x40 Screw"] = ngProcess
                    excelData["Process 4 M4x40 Screw Lot No"] = ngProcess
                    excelData["Process 4 PartitionGasket"] = ngProcess
                    excelData["Process 4 PartitionGasket Lot No"] = ngProcess
                    excelData["Process 4 M4x12 Screw"] = ngProcess
                    excelData["Process 4 M4x12 Screw Lot No"] = ngProcess
                    excelData["Process 4 Muffler"] = ngProcess
                    excelData["Process 4 Muffler Lot No"] = ngProcess
                    excelData["Process 4 Muffler Gasket"] = ngProcess
                    excelData["Process 4 Muffler Gasket Lot No"] = ngProcess
                    excelData["Process 4 VCR"] = ngProcess
                    excelData["Process 4 VCR Lot No"] = ngProcess
                    excelData["Process 4 ST"] = ngProcess
                    excelData["Process 4 Actual Time"] = ngProcess
                    excelData["Process 4 NG Cause"] = ngProcess
                    excelData["Process 4 Repaired Action"] = ngProcess

                    excelData["Process 5 Model Code"] = ngProcess
                    excelData["Process 5 S/N"] = ngProcess
                    excelData["Process 5 ID"] = ngProcess
                    excelData["Process 5 NAME"] = ngProcess
                    excelData["Process 5 Regular/Contractual"] = ngProcess
                    excelData["Process 5 Rating Label"] = ngProcess
                    excelData["Process 5 Rating Label Lot No"] = ngProcess
                    excelData["Process 5 ST"] = ngProcess
                    excelData["Process 5 Actual Time"] = ngProcess
                    excelData["Process 5 NG Cause"] = ngProcess
                    excelData["Process 5 Repaired Action"] = ngProcess

                    excelData["Process 6 Model Code"] = ngProcess
                    excelData["Process 6 S/N"] = ngProcess
                    excelData["Process 6 ID"] = ngProcess
                    excelData["Process 6 NAME"] = ngProcess
                    excelData["Process 6 Regular/Contractual"] = ngProcess
                    excelData["Process 6 Vinyl"] = ngProcess
                    excelData["Process 6 Vinyl Lot No"] = ngProcess
                    excelData["Process 6 ST"] = ngProcess
                    excelData["Process 6 Actual Time"] = ngProcess
                    excelData["Process 6 NG Cause"] = ngProcess
                    excelData["Process 6 Repaired Action"] = ngProcess
                    
                elif process4Status == "NG":
                    repairedProcess = "REPAIRED AT PROCESS2"
                    ngProcess = "NG AT PROCESS4"
                    process2Row += 1
                    process3Row += 1
                    process4Row += 1

                    excelData["Process 1 Model Code"] = repairedProcess
                    excelData["Process 1 S/N"] = repairedProcess
                    excelData["Process 1 ID"] = repairedProcess
                    excelData["Process 1 NAME"] = repairedProcess
                    excelData["Process 1 Regular/Contractual"] = repairedProcess
                    excelData["Process 1 Em2p"] = repairedProcess
                    excelData["Process 1 Em2p Lot No"] = repairedProcess
                    excelData["Process 1 Em3p"] = repairedProcess
                    excelData["Process 1 Em3p Lot No"] = repairedProcess
                    excelData["Process 1 Harness"] = repairedProcess
                    excelData["Process 1 Harness Lot No"] = repairedProcess
                    excelData["Process 1 Frame"] = repairedProcess
                    excelData["Process 1 Frame Lot No"] = repairedProcess 
                    excelData["Process 1 Bushing"] = repairedProcess
                    excelData["Process 1 Bushing Lot No"] = repairedProcess
                    excelData["Process 1 ST"] = repairedProcess
                    excelData["Process 1 Actual Time"] = repairedProcess
                    excelData["Process 1 NG Cause"] = repairedProcess
                    excelData["Process 1 Repaired Action"] = repairedProcess 

                    excelData["Process 5 Model Code"] = ngProcess
                    excelData["Process 5 S/N"] = ngProcess
                    excelData["Process 5 ID"] = ngProcess
                    excelData["Process 5 NAME"] = ngProcess
                    excelData["Process 5 Regular/Contractual"] = ngProcess
                    excelData["Process 5 Rating Label"] = ngProcess
                    excelData["Process 5 Rating Label Lot No"] = ngProcess
                    excelData["Process 5 ST"] = ngProcess
                    excelData["Process 5 Actual Time"] = ngProcess
                    excelData["Process 5 NG Cause"] = ngProcess
                    excelData["Process 5 Repaired Action"] = ngProcess

                    excelData["Process 6 Model Code"] = ngProcess
                    excelData["Process 6 S/N"] = ngProcess
                    excelData["Process 6 ID"] = ngProcess
                    excelData["Process 6 NAME"] = ngProcess
                    excelData["Process 6 Regular/Contractual"] = ngProcess
                    excelData["Process 6 Vinyl"] = ngProcess
                    excelData["Process 6 Vinyl Lot No"] = ngProcess
                    excelData["Process 6 ST"] = ngProcess
                    excelData["Process 6 Actual Time"] = ngProcess
                    excelData["Process 6 NG Cause"] = ngProcess
                    excelData["Process 6 Repaired Action"] = ngProcess

                elif process5Status == "NG PRESSURE":
                    repairedProcess = "REPAIRED AT PROCESS2"
                    ngProcess = "NG PRESSURE AT PROCESS5"
                    process2Row += 1
                    process3Row += 1
                    process4Row += 1
                    process5Row += 1
                    piRow += 1

                    excelData["DATE"] = ngProcess
                    excelData["TIME"] = ngProcess
                    excelData["MODEL CODE"] = ngProcess
                    # excelData["PROCESS S/N"] = ngProcess
                    excelData["S/N"] = ngProcess
                    excelData["PASS/NG"] = ngProcess
                    excelData["VOLTAGE MAX (V)"] = ngProcess
                    excelData["WATTAGE MAX (W)"] = ngProcess
                    excelData["CLOSED PRESSURE_MAX (kPa)"] = ngProcess
                    excelData["VOLTAGE Middle (V)"] = ngProcess
                    excelData["WATTAGE Middle (W)"] = ngProcess
                    excelData["AMPERAGE Middle (A)"] = ngProcess
                    excelData["CLOSED PRESSURE Middle (kPa)"] = ngProcess
                    excelData["dB(A) 1"] = ngProcess
                    excelData["dB(A) 2"] = ngProcess
                    excelData["dB(A) 3"] = ngProcess
                    excelData["VOLTAGE MIN (V)"] = ngProcess
                    excelData["WATTAGE MIN (W)"] = ngProcess
                    excelData["CLOSED PRESSURE MIN (kPa)"] = ngProcess

                    excelData["Process 1 Model Code"] = repairedProcess
                    excelData["Process 1 S/N"] = repairedProcess
                    excelData["Process 1 ID"] = repairedProcess
                    excelData["Process 1 NAME"] = repairedProcess
                    excelData["Process 1 Regular/Contractual"] = repairedProcess
                    excelData["Process 1 Em2p"] = repairedProcess
                    excelData["Process 1 Em2p Lot No"] = repairedProcess
                    excelData["Process 1 Em3p"] = repairedProcess
                    excelData["Process 1 Em3p Lot No"] = repairedProcess
                    excelData["Process 1 Harness"] = repairedProcess
                    excelData["Process 1 Harness Lot No"] = repairedProcess
                    excelData["Process 1 Frame"] = repairedProcess
                    excelData["Process 1 Frame Lot No"] = repairedProcess 
                    excelData["Process 1 Bushing"] = repairedProcess
                    excelData["Process 1 Bushing Lot No"] = repairedProcess
                    excelData["Process 1 ST"] = repairedProcess
                    excelData["Process 1 Actual Time"] = repairedProcess
                    excelData["Process 1 NG Cause"] = repairedProcess
                    excelData["Process 1 Repaired Action"] = repairedProcess  

                    excelData["Process 6 Model Code"] = ngProcess
                    excelData["Process 6 S/N"] = ngProcess
                    excelData["Process 6 ID"] = ngProcess
                    excelData["Process 6 NAME"] = ngProcess
                    excelData["Process 6 Regular/Contractual"] = ngProcess
                    excelData["Process 6 Vinyl"] = ngProcess
                    excelData["Process 6 Vinyl Lot No"] = ngProcess
                    excelData["Process 6 ST"] = ngProcess
                    excelData["Process 6 Actual Time"] = ngProcess
                    excelData["Process 6 NG Cause"] = ngProcess
                    excelData["Process 6 Repaired Action"] = ngProcess

                elif process5Status == "NG":
                    repairedProcess = "REPAIRED AT PROCESS2"
                    ngProcess = "NG AT PROCESS5"
                    process2Row += 1
                    process3Row += 1
                    process4Row += 1
                    process5Row += 1
                    piRow += 1

                    excelData["Process 1 Model Code"] = repairedProcess
                    excelData["Process 1 S/N"] = repairedProcess
                    excelData["Process 1 ID"] = repairedProcess
                    excelData["Process 1 NAME"] = repairedProcess
                    excelData["Process 1 Regular/Contractual"] = repairedProcess
                    excelData["Process 1 Em2p"] = repairedProcess
                    excelData["Process 1 Em2p Lot No"] = repairedProcess
                    excelData["Process 1 Em3p"] = repairedProcess
                    excelData["Process 1 Em3p Lot No"] = repairedProcess
                    excelData["Process 1 Harness"] = repairedProcess
                    excelData["Process 1 Harness Lot No"] = repairedProcess
                    excelData["Process 1 Frame"] = repairedProcess
                    excelData["Process 1 Frame Lot No"] = repairedProcess 
                    excelData["Process 1 Bushing"] = repairedProcess
                    excelData["Process 1 Bushing Lot No"] = repairedProcess
                    excelData["Process 1 ST"] = repairedProcess
                    excelData["Process 1 Actual Time"] = repairedProcess
                    excelData["Process 1 NG Cause"] = repairedProcess
                    excelData["Process 1 Repaired Action"] = repairedProcess  

                    excelData["Process 6 Model Code"] = ngProcess
                    excelData["Process 6 S/N"] = ngProcess
                    excelData["Process 6 ID"] = ngProcess
                    excelData["Process 6 NAME"] = ngProcess
                    excelData["Process 6 Regular/Contractual"] = ngProcess
                    excelData["Process 6 Vinyl"] = ngProcess
                    excelData["Process 6 Vinyl Lot No"] = ngProcess
                    excelData["Process 6 ST"] = ngProcess
                    excelData["Process 6 Actual Time"] = ngProcess
                    excelData["Process 6 NG Cause"] = ngProcess
                    excelData["Process 6 Repaired Action"] = ngProcess

                elif process6Status == "NG":
                    repairedProcess = "REPAIRED AT PROCESS2"
                    
                    process2Row += 1
                    process3Row += 1
                    process4Row += 1
                    process5Row += 1
                    process6Row += 1
                    piRow += 1

                    excelData["Process 1 Model Code"] = repairedProcess
                    excelData["Process 1 S/N"] = repairedProcess
                    excelData["Process 1 ID"] = repairedProcess
                    excelData["Process 1 NAME"] = repairedProcess
                    excelData["Process 1 Regular/Contractual"] = repairedProcess
                    excelData["Process 1 Em2p"] = repairedProcess
                    excelData["Process 1 Em2p Lot No"] = repairedProcess
                    excelData["Process 1 Em3p"] = repairedProcess
                    excelData["Process 1 Em3p Lot No"] = repairedProcess
                    excelData["Process 1 Harness"] = repairedProcess
                    excelData["Process 1 Harness Lot No"] = repairedProcess
                    excelData["Process 1 Frame"] = repairedProcess
                    excelData["Process 1 Frame Lot No"] = repairedProcess 
                    excelData["Process 1 Bushing"] = repairedProcess
                    excelData["Process 1 Bushing Lot No"] = repairedProcess
                    excelData["Process 1 ST"] = repairedProcess
                    excelData["Process 1 Actual Time"] = repairedProcess
                    excelData["Process 1 NG Cause"] = repairedProcess
                    excelData["Process 1 Repaired Action"] = repairedProcess  

            elif process3Status == "Repaired":
                if process4Status == "NG":
                    repairedProcess = "REPAIRED AT PROCESS3"
                    ngProcess = "NG AT PROCESS4"
                    process3Row += 1
                    process4Row += 1

                    excelData["Process 1 Model Code"] = repairedProcess
                    excelData["Process 1 S/N"] = repairedProcess
                    excelData["Process 1 ID"] = repairedProcess
                    excelData["Process 1 NAME"] = repairedProcess
                    excelData["Process 1 Regular/Contractual"] = repairedProcess
                    excelData["Process 1 Em2p"] = repairedProcess
                    excelData["Process 1 Em2p Lot No"] = repairedProcess
                    excelData["Process 1 Em3p"] = repairedProcess
                    excelData["Process 1 Em3p Lot No"] = repairedProcess
                    excelData["Process 1 Harness"] = repairedProcess
                    excelData["Process 1 Harness Lot No"] = repairedProcess
                    excelData["Process 1 Frame"] = repairedProcess
                    excelData["Process 1 Frame Lot No"] = repairedProcess 
                    excelData["Process 1 Bushing"] = repairedProcess
                    excelData["Process 1 Bushing Lot No"] = repairedProcess
                    excelData["Process 1 ST"] = repairedProcess
                    excelData["Process 1 Actual Time"] = repairedProcess
                    excelData["Process 1 NG Cause"] = repairedProcess
                    excelData["Process 1 Repaired Action"] = repairedProcess  

                    excelData["Process 2 Model Code"] = repairedProcess
                    excelData["Process 2 S/N"] = repairedProcess
                    excelData["Process 2 ID"] = repairedProcess
                    excelData["Process 2 NAME"] = repairedProcess
                    excelData["Process 2 Regular/Contractual"] = repairedProcess
                    excelData["Process 2 M4x40 Screw"] = repairedProcess
                    excelData["Process 2 M4x40 Screw Lot No"] = repairedProcess
                    excelData["Process 2 Rod Blk"] = repairedProcess
                    excelData["Process 2 Rod Blk Lot No"] = repairedProcess
                    excelData["Process 2 Df Blk"] = repairedProcess
                    excelData["Process 2 Df Blk Lot No"] = repairedProcess
                    excelData["Process 2 Df Ring"] = repairedProcess
                    excelData["Process 2 Df Ring Lot No"] = repairedProcess
                    excelData["Process 2 Washer"] = repairedProcess
                    excelData["Process 2 Washer Lot No"] = repairedProcess
                    excelData["Process 2 Lock Nut"] = repairedProcess
                    excelData["Process 2 Lock Nut Lot No"] = repairedProcess
                    excelData["Process 2 ST"] = repairedProcess
                    excelData["Process 2 Actual Time"] = repairedProcess
                    excelData["Process 2 NG Cause"] = repairedProcess
                    excelData["Process 2 Repaired Action"] = repairedProcess

                    excelData["Process 5 Model Code"] = ngProcess
                    excelData["Process 5 S/N"] = ngProcess
                    excelData["Process 5 ID"] = ngProcess
                    excelData["Process 5 NAME"] = ngProcess
                    excelData["Process 5 Regular/Contractual"] = ngProcess
                    excelData["Process 5 Rating Label"] = ngProcess
                    excelData["Process 5 Rating Label Lot No"] = ngProcess
                    excelData["Process 5 ST"] = ngProcess
                    excelData["Process 5 Actual Time"] = ngProcess
                    excelData["Process 5 NG Cause"] = ngProcess
                    excelData["Process 5 Repaired Action"] = ngProcess

                    excelData["Process 6 Model Code"] = ngProcess
                    excelData["Process 6 S/N"] = ngProcess
                    excelData["Process 6 ID"] = ngProcess
                    excelData["Process 6 NAME"] = ngProcess
                    excelData["Process 6 Regular/Contractual"] = ngProcess
                    excelData["Process 6 Vinyl"] = ngProcess
                    excelData["Process 6 Vinyl Lot No"] = ngProcess
                    excelData["Process 6 ST"] = ngProcess
                    excelData["Process 6 Actual Time"] = ngProcess
                    excelData["Process 6 NG Cause"] = ngProcess
                    excelData["Process 6 Repaired Action"] = ngProcess

                elif process5Status == "NG PRESSURE":
                    repairedProcess = "REPAIRED AT PROCESS3"
                    ngProcess = "NG PRESSURE AT PROCESS5"
                    process3Row += 1
                    process4Row += 1
                    process5Row += 1
                    piRow += 1

                    excelData["DATE"] = ngProcess
                    excelData["TIME"] = ngProcess
                    excelData["MODEL CODE"] = ngProcess
                    # excelData["PROCESS S/N"] = ngProcess
                    excelData["S/N"] = ngProcess
                    excelData["PASS/NG"] = ngProcess
                    excelData["VOLTAGE MAX (V)"] = ngProcess
                    excelData["WATTAGE MAX (W)"] = ngProcess
                    excelData["CLOSED PRESSURE_MAX (kPa)"] = ngProcess
                    excelData["VOLTAGE Middle (V)"] = ngProcess
                    excelData["WATTAGE Middle (W)"] = ngProcess
                    excelData["AMPERAGE Middle (A)"] = ngProcess
                    excelData["CLOSED PRESSURE Middle (kPa)"] = ngProcess
                    excelData["dB(A) 1"] = ngProcess
                    excelData["dB(A) 2"] = ngProcess
                    excelData["dB(A) 3"] = ngProcess
                    excelData["VOLTAGE MIN (V)"] = ngProcess
                    excelData["WATTAGE MIN (W)"] = ngProcess
                    excelData["CLOSED PRESSURE MIN (kPa)"] = ngProcess

                    excelData["Process 1 Model Code"] = repairedProcess
                    excelData["Process 1 S/N"] = repairedProcess
                    excelData["Process 1 ID"] = repairedProcess
                    excelData["Process 1 NAME"] = repairedProcess
                    excelData["Process 1 Regular/Contractual"] = repairedProcess
                    excelData["Process 1 Em2p"] = repairedProcess
                    excelData["Process 1 Em2p Lot No"] = repairedProcess
                    excelData["Process 1 Em3p"] = repairedProcess
                    excelData["Process 1 Em3p Lot No"] = repairedProcess
                    excelData["Process 1 Harness"] = repairedProcess
                    excelData["Process 1 Harness Lot No"] = repairedProcess
                    excelData["Process 1 Frame"] = repairedProcess
                    excelData["Process 1 Frame Lot No"] = repairedProcess 
                    excelData["Process 1 Bushing"] = repairedProcess
                    excelData["Process 1 Bushing Lot No"] = repairedProcess
                    excelData["Process 1 ST"] = repairedProcess
                    excelData["Process 1 Actual Time"] = repairedProcess
                    excelData["Process 1 NG Cause"] = repairedProcess
                    excelData["Process 1 Repaired Action"] = repairedProcess  

                    excelData["Process 2 Model Code"] = repairedProcess
                    excelData["Process 2 S/N"] = repairedProcess
                    excelData["Process 2 ID"] = repairedProcess
                    excelData["Process 2 NAME"] = repairedProcess
                    excelData["Process 2 Regular/Contractual"] = repairedProcess
                    excelData["Process 2 M4x40 Screw"] = repairedProcess
                    excelData["Process 2 M4x40 Screw Lot No"] = repairedProcess
                    excelData["Process 2 Rod Blk"] = repairedProcess
                    excelData["Process 2 Rod Blk Lot No"] = repairedProcess
                    excelData["Process 2 Df Blk"] = repairedProcess
                    excelData["Process 2 Df Blk Lot No"] = repairedProcess
                    excelData["Process 2 Df Ring"] = repairedProcess
                    excelData["Process 2 Df Ring Lot No"] = repairedProcess
                    excelData["Process 2 Washer"] = repairedProcess
                    excelData["Process 2 Washer Lot No"] = repairedProcess
                    excelData["Process 2 Lock Nut"] = repairedProcess
                    excelData["Process 2 Lock Nut Lot No"] = repairedProcess
                    excelData["Process 2 ST"] = repairedProcess
                    excelData["Process 2 Actual Time"] = repairedProcess
                    excelData["Process 2 NG Cause"] = repairedProcess
                    excelData["Process 2 Repaired Action"] = repairedProcess

                    excelData["Process 6 Model Code"] = ngProcess
                    excelData["Process 6 S/N"] = ngProcess
                    excelData["Process 6 ID"] = ngProcess
                    excelData["Process 6 NAME"] = ngProcess
                    excelData["Process 6 Regular/Contractual"] = ngProcess
                    excelData["Process 6 Vinyl"] = ngProcess
                    excelData["Process 6 Vinyl Lot No"] = ngProcess
                    excelData["Process 6 ST"] = ngProcess
                    excelData["Process 6 Actual Time"] = ngProcess
                    excelData["Process 6 NG Cause"] = ngProcess
                    excelData["Process 6 Repaired Action"] = ngProcess

                elif process5Status == "NG":
                    repairedProcess = "REPAIRED AT PROCESS3"
                    ngProcess = "NG AT PROCESS5"
                    process3Row += 1
                    process4Row += 1
                    process5Row += 1
                    piRow += 1

                    excelData["Process 1 Model Code"] = repairedProcess
                    excelData["Process 1 S/N"] = repairedProcess
                    excelData["Process 1 ID"] = repairedProcess
                    excelData["Process 1 NAME"] = repairedProcess
                    excelData["Process 1 Regular/Contractual"] = repairedProcess
                    excelData["Process 1 Em2p"] = repairedProcess
                    excelData["Process 1 Em2p Lot No"] = repairedProcess
                    excelData["Process 1 Em3p"] = repairedProcess
                    excelData["Process 1 Em3p Lot No"] = repairedProcess
                    excelData["Process 1 Harness"] = repairedProcess
                    excelData["Process 1 Harness Lot No"] = repairedProcess
                    excelData["Process 1 Frame"] = repairedProcess
                    excelData["Process 1 Frame Lot No"] = repairedProcess 
                    excelData["Process 1 Bushing"] = repairedProcess
                    excelData["Process 1 Bushing Lot No"] = repairedProcess
                    excelData["Process 1 ST"] = repairedProcess
                    excelData["Process 1 Actual Time"] = repairedProcess
                    excelData["Process 1 NG Cause"] = repairedProcess
                    excelData["Process 1 Repaired Action"] = repairedProcess  

                    excelData["Process 2 Model Code"] = repairedProcess
                    excelData["Process 2 S/N"] = repairedProcess
                    excelData["Process 2 ID"] = repairedProcess
                    excelData["Process 2 NAME"] = repairedProcess
                    excelData["Process 2 Regular/Contractual"] = repairedProcess
                    excelData["Process 2 M4x40 Screw"] = repairedProcess
                    excelData["Process 2 M4x40 Screw Lot No"] = repairedProcess
                    excelData["Process 2 Rod Blk"] = repairedProcess
                    excelData["Process 2 Rod Blk Lot No"] = repairedProcess
                    excelData["Process 2 Df Blk"] = repairedProcess
                    excelData["Process 2 Df Blk Lot No"] = repairedProcess
                    excelData["Process 2 Df Ring"] = repairedProcess
                    excelData["Process 2 Df Ring Lot No"] = repairedProcess
                    excelData["Process 2 Washer"] = repairedProcess
                    excelData["Process 2 Washer Lot No"] = repairedProcess
                    excelData["Process 2 Lock Nut"] = repairedProcess
                    excelData["Process 2 Lock Nut Lot No"] = repairedProcess
                    excelData["Process 2 ST"] = repairedProcess
                    excelData["Process 2 Actual Time"] = repairedProcess
                    excelData["Process 2 NG Cause"] = repairedProcess
                    excelData["Process 2 Repaired Action"] = repairedProcess

                    excelData["Process 6 Model Code"] = ngProcess
                    excelData["Process 6 S/N"] = ngProcess
                    excelData["Process 6 ID"] = ngProcess
                    excelData["Process 6 NAME"] = ngProcess
                    excelData["Process 6 Regular/Contractual"] = ngProcess
                    excelData["Process 6 Vinyl"] = ngProcess
                    excelData["Process 6 Vinyl Lot No"] = ngProcess
                    excelData["Process 6 ST"] = ngProcess
                    excelData["Process 6 Actual Time"] = ngProcess
                    excelData["Process 6 NG Cause"] = ngProcess
                    excelData["Process 6 Repaired Action"] = ngProcess

                elif process6Status == "NG":
                    repairedProcess = "REPAIRED AT PROCESS3"
                    ngProcess = "NG AT PROCESS6"
                    process3Row += 1
                    process4Row += 1
                    process5Row += 1
                    process6Row += 1
                    piRow += 1

                    excelData["Process 1 Model Code"] = repairedProcess
                    excelData["Process 1 S/N"] = repairedProcess
                    excelData["Process 1 ID"] = repairedProcess
                    excelData["Process 1 NAME"] = repairedProcess
                    excelData["Process 1 Regular/Contractual"] = repairedProcess
                    excelData["Process 1 Em2p"] = repairedProcess
                    excelData["Process 1 Em2p Lot No"] = repairedProcess
                    excelData["Process 1 Em3p"] = repairedProcess
                    excelData["Process 1 Em3p Lot No"] = repairedProcess
                    excelData["Process 1 Harness"] = repairedProcess
                    excelData["Process 1 Harness Lot No"] = repairedProcess
                    excelData["Process 1 Frame"] = repairedProcess
                    excelData["Process 1 Frame Lot No"] = repairedProcess 
                    excelData["Process 1 Bushing"] = repairedProcess
                    excelData["Process 1 Bushing Lot No"] = repairedProcess
                    excelData["Process 1 ST"] = repairedProcess
                    excelData["Process 1 Actual Time"] = repairedProcess
                    excelData["Process 1 NG Cause"] = repairedProcess
                    excelData["Process 1 Repaired Action"] = repairedProcess  

                    excelData["Process 2 Model Code"] = repairedProcess
                    excelData["Process 2 S/N"] = repairedProcess
                    excelData["Process 2 ID"] = repairedProcess
                    excelData["Process 2 NAME"] = repairedProcess
                    excelData["Process 2 Regular/Contractual"] = repairedProcess
                    excelData["Process 2 M4x40 Screw"] = repairedProcess
                    excelData["Process 2 M4x40 Screw Lot No"] = repairedProcess
                    excelData["Process 2 Rod Blk"] = repairedProcess
                    excelData["Process 2 Rod Blk Lot No"] = repairedProcess
                    excelData["Process 2 Df Blk"] = repairedProcess
                    excelData["Process 2 Df Blk Lot No"] = repairedProcess
                    excelData["Process 2 Df Ring"] = repairedProcess
                    excelData["Process 2 Df Ring Lot No"] = repairedProcess
                    excelData["Process 2 Washer"] = repairedProcess
                    excelData["Process 2 Washer Lot No"] = repairedProcess
                    excelData["Process 2 Lock Nut"] = repairedProcess
                    excelData["Process 2 Lock Nut Lot No"] = repairedProcess
                    excelData["Process 2 ST"] = repairedProcess
                    excelData["Process 2 Actual Time"] = repairedProcess
                    excelData["Process 2 NG Cause"] = repairedProcess
                    excelData["Process 2 Repaired Action"] = repairedProcess

            elif process4Status == "Repaired":
                if process5Status == "NG PRESSURE":
                    repairedProcess = "REPAIRED AT PROCESS4"
                    ngProcess = "NG PRESSURE AT PROCESS5"
                    process4Row += 1
                    process5Row += 1
                    piRow += 1

                    excelData["DATE"] = ngProcess
                    excelData["TIME"] = ngProcess
                    excelData["MODEL CODE"] = ngProcess
                    # excelData["PROCESS S/N"] = ngProcess
                    excelData["S/N"] = ngProcess
                    excelData["PASS/NG"] = ngProcess
                    excelData["VOLTAGE MAX (V)"] = ngProcess
                    excelData["WATTAGE MAX (W)"] = ngProcess
                    excelData["CLOSED PRESSURE_MAX (kPa)"] = ngProcess
                    excelData["VOLTAGE Middle (V)"] = ngProcess
                    excelData["WATTAGE Middle (W)"] = ngProcess
                    excelData["AMPERAGE Middle (A)"] = ngProcess
                    excelData["CLOSED PRESSURE Middle (kPa)"] = ngProcess
                    excelData["dB(A) 1"] = ngProcess
                    excelData["dB(A) 2"] = ngProcess
                    excelData["dB(A) 3"] = ngProcess
                    excelData["VOLTAGE MIN (V)"] = ngProcess
                    excelData["WATTAGE MIN (W)"] = ngProcess
                    excelData["CLOSED PRESSURE MIN (kPa)"] = ngProcess

                    excelData["Process 1 Model Code"] = repairedProcess
                    excelData["Process 1 S/N"] = repairedProcess
                    excelData["Process 1 ID"] = repairedProcess
                    excelData["Process 1 NAME"] = repairedProcess
                    excelData["Process 1 Regular/Contractual"] = repairedProcess
                    excelData["Process 1 Em2p"] = repairedProcess
                    excelData["Process 1 Em2p Lot No"] = repairedProcess
                    excelData["Process 1 Em3p"] = repairedProcess
                    excelData["Process 1 Em3p Lot No"] = repairedProcess
                    excelData["Process 1 Harness"] = repairedProcess
                    excelData["Process 1 Harness Lot No"] = repairedProcess
                    excelData["Process 1 Frame"] = repairedProcess
                    excelData["Process 1 Frame Lot No"] = repairedProcess 
                    excelData["Process 1 Bushing"] = repairedProcess
                    excelData["Process 1 Bushing Lot No"] = repairedProcess
                    excelData["Process 1 ST"] = repairedProcess
                    excelData["Process 1 Actual Time"] = repairedProcess
                    excelData["Process 1 NG Cause"] = repairedProcess
                    excelData["Process 1 Repaired Action"] = repairedProcess  

                    excelData["Process 2 Model Code"] = repairedProcess
                    excelData["Process 2 S/N"] = repairedProcess
                    excelData["Process 2 ID"] = repairedProcess
                    excelData["Process 2 NAME"] = repairedProcess
                    excelData["Process 2 Regular/Contractual"] = repairedProcess
                    excelData["Process 2 M4x40 Screw"] = repairedProcess
                    excelData["Process 2 M4x40 Screw Lot No"] = repairedProcess
                    excelData["Process 2 Rod Blk"] = repairedProcess
                    excelData["Process 2 Rod Blk Lot No"] = repairedProcess
                    excelData["Process 2 Df Blk"] = repairedProcess
                    excelData["Process 2 Df Blk Lot No"] = repairedProcess
                    excelData["Process 2 Df Ring"] = repairedProcess
                    excelData["Process 2 Df Ring Lot No"] = repairedProcess
                    excelData["Process 2 Washer"] = repairedProcess
                    excelData["Process 2 Washer Lot No"] = repairedProcess
                    excelData["Process 2 Lock Nut"] = repairedProcess
                    excelData["Process 2 Lock Nut Lot No"] = repairedProcess
                    excelData["Process 2 ST"] = repairedProcess
                    excelData["Process 2 Actual Time"] = repairedProcess
                    excelData["Process 2 NG Cause"] = repairedProcess
                    excelData["Process 2 Repaired Action"] = repairedProcess

                    excelData["Process 3 Model Code"] = ngProcess
                    excelData["Process 3 S/N"] = ngProcess
                    excelData["Process 3 ID"] = ngProcess
                    excelData["Process 3 NAME"] = ngProcess
                    excelData["Process 3 Regular/Contractual"] = ngProcess
                    excelData["Process 3 Frame Gasket"] = ngProcess
                    excelData["Process 3 Frame Gasket Lot No"] = ngProcess
                    excelData["Process 3 Casing Block"] = ngProcess
                    excelData["Process 3 Casing Block Lot No"] = ngProcess
                    excelData["Process 3 Casing Block Inspection 1 Average Data"] = ngProcess
                    excelData["Process 3 Casing Block Inspection 1 Minimum Data"] = ngProcess
                    excelData["Process 3 Casing Block Inspection 1 Maximum Data"] = ngProcess
                    excelData["Process 3 Casing Gasket"] = ngProcess
                    excelData["Process 3 Casing Gasket Lot No"] = ngProcess
                    excelData["Process 3 M4x16 Screw 1"] = ngProcess
                    excelData["Process 3 M4x16 Screw 1 Lot No"] = ngProcess
                    excelData["Process 3 M4x16 Screw 2"] = ngProcess
                    excelData["Process 3 M4x16 Screw 2 Lot No"] = ngProcess
                    excelData["Process 3 Ball Cushion"] = ngProcess
                    excelData["Process 3 Ball Cushion Lot No"] = ngProcess
                    excelData["Process 3 Frame Cover"] = ngProcess
                    excelData["Process 3 Frame Cover Lot No"] = ngProcess
                    excelData["Process 3 Partition Board"] = ngProcess
                    excelData["Process 3 Partition Board Lot No"] = ngProcess
                    excelData["Process 3 Built In Tube 1"] = ngProcess
                    excelData["Process 3 Built In Tube 1 Lot No"] = ngProcess
                    excelData["Process 3 Built In Tube 2"] = ngProcess
                    excelData["Process 3 Built In Tube 2 Lot No"] = ngProcess
                    excelData["Process 3 Head Cover"] = ngProcess
                    excelData["Process 3 Head Cover Lot No"] = ngProcess
                    excelData["Process 3 Casing Packing"] = ngProcess
                    excelData["Process 3 Casing Packing Lot No"] = ngProcess
                    excelData["Process 3 M4x12 Screw"] = ngProcess
                    excelData["Process 3 M4x12 Screw Lot No"] = ngProcess
                    excelData["Process 3 Csb L"] = ngProcess
                    excelData["Process 3 Csb L Lot No"] = ngProcess
                    excelData["Process 3 Csb R"] = ngProcess
                    excelData["Process 3 Csb R Lot No"] = ngProcess
                    excelData["Process 3 Head Packing"] = ngProcess
                    excelData["Process 3 Head Packing Lot No"] = ngProcess
                    excelData["Process 3 ST"] = ngProcess
                    excelData["Process 3 Actual Time"] = ngProcess
                    excelData["Process 3 NG Cause"] = ngProcess
                    excelData["Process 3 Repaired Action"] = ngProcess

                    excelData["Process 6 Model Code"] = ngProcess
                    excelData["Process 6 S/N"] = ngProcess
                    excelData["Process 6 ID"] = ngProcess
                    excelData["Process 6 NAME"] = ngProcess
                    excelData["Process 6 Regular/Contractual"] = ngProcess
                    excelData["Process 6 Vinyl"] = ngProcess
                    excelData["Process 6 Vinyl Lot No"] = ngProcess
                    excelData["Process 6 ST"] = ngProcess
                    excelData["Process 6 Actual Time"] = ngProcess
                    excelData["Process 6 NG Cause"] = ngProcess
                    excelData["Process 6 Repaired Action"] = ngProcess

                elif process5Status == "NG":
                    repairedProcess = "REPAIRED AT PROCESS4"
                    ngProcess = "NG AT PROCESS5"
                    process4Row += 1
                    process5Row += 1
                    piRow += 1

                    excelData["Process 1 Model Code"] = repairedProcess
                    excelData["Process 1 S/N"] = repairedProcess
                    excelData["Process 1 ID"] = repairedProcess
                    excelData["Process 1 NAME"] = repairedProcess
                    excelData["Process 1 Regular/Contractual"] = repairedProcess
                    excelData["Process 1 Em2p"] = repairedProcess
                    excelData["Process 1 Em2p Lot No"] = repairedProcess
                    excelData["Process 1 Em3p"] = repairedProcess
                    excelData["Process 1 Em3p Lot No"] = repairedProcess
                    excelData["Process 1 Harness"] = repairedProcess
                    excelData["Process 1 Harness Lot No"] = repairedProcess
                    excelData["Process 1 Frame"] = repairedProcess
                    excelData["Process 1 Frame Lot No"] = repairedProcess 
                    excelData["Process 1 Bushing"] = repairedProcess
                    excelData["Process 1 Bushing Lot No"] = repairedProcess
                    excelData["Process 1 ST"] = repairedProcess
                    excelData["Process 1 Actual Time"] = repairedProcess
                    excelData["Process 1 NG Cause"] = repairedProcess
                    excelData["Process 1 Repaired Action"] = repairedProcess  

                    excelData["Process 2 Model Code"] = repairedProcess
                    excelData["Process 2 S/N"] = repairedProcess
                    excelData["Process 2 ID"] = repairedProcess
                    excelData["Process 2 NAME"] = repairedProcess
                    excelData["Process 2 Regular/Contractual"] = repairedProcess
                    excelData["Process 2 M4x40 Screw"] = repairedProcess
                    excelData["Process 2 M4x40 Screw Lot No"] = repairedProcess
                    excelData["Process 2 Rod Blk"] = repairedProcess
                    excelData["Process 2 Rod Blk Lot No"] = repairedProcess
                    excelData["Process 2 Df Blk"] = repairedProcess
                    excelData["Process 2 Df Blk Lot No"] = repairedProcess
                    excelData["Process 2 Df Ring"] = repairedProcess
                    excelData["Process 2 Df Ring Lot No"] = repairedProcess
                    excelData["Process 2 Washer"] = repairedProcess
                    excelData["Process 2 Washer Lot No"] = repairedProcess
                    excelData["Process 2 Lock Nut"] = repairedProcess
                    excelData["Process 2 Lock Nut Lot No"] = repairedProcess
                    excelData["Process 2 ST"] = repairedProcess
                    excelData["Process 2 Actual Time"] = repairedProcess
                    excelData["Process 2 NG Cause"] = repairedProcess
                    excelData["Process 2 Repaired Action"] = repairedProcess

                    excelData["Process 3 Model Code"] = ngProcess
                    excelData["Process 3 S/N"] = ngProcess
                    excelData["Process 3 ID"] = ngProcess
                    excelData["Process 3 NAME"] = ngProcess
                    excelData["Process 3 Regular/Contractual"] = ngProcess
                    excelData["Process 3 Frame Gasket"] = ngProcess
                    excelData["Process 3 Frame Gasket Lot No"] = ngProcess
                    excelData["Process 3 Casing Block"] = ngProcess
                    excelData["Process 3 Casing Block Lot No"] = ngProcess
                    excelData["Process 3 Casing Block Inspection 1 Average Data"] = ngProcess
                    excelData["Process 3 Casing Block Inspection 1 Minimum Data"] = ngProcess
                    excelData["Process 3 Casing Block Inspection 1 Maximum Data"] = ngProcess
                    excelData["Process 3 Casing Gasket"] = ngProcess
                    excelData["Process 3 Casing Gasket Lot No"] = ngProcess
                    excelData["Process 3 M4x16 Screw 1"] = ngProcess
                    excelData["Process 3 M4x16 Screw 1 Lot No"] = ngProcess
                    excelData["Process 3 M4x16 Screw 2"] = ngProcess
                    excelData["Process 3 M4x16 Screw 2 Lot No"] = ngProcess
                    excelData["Process 3 Ball Cushion"] = ngProcess
                    excelData["Process 3 Ball Cushion Lot No"] = ngProcess
                    excelData["Process 3 Frame Cover"] = ngProcess
                    excelData["Process 3 Frame Cover Lot No"] = ngProcess
                    excelData["Process 3 Partition Board"] = ngProcess
                    excelData["Process 3 Partition Board Lot No"] = ngProcess
                    excelData["Process 3 Built In Tube 1"] = ngProcess
                    excelData["Process 3 Built In Tube 1 Lot No"] = ngProcess
                    excelData["Process 3 Built In Tube 2"] = ngProcess
                    excelData["Process 3 Built In Tube 2 Lot No"] = ngProcess
                    excelData["Process 3 Head Cover"] = ngProcess
                    excelData["Process 3 Head Cover Lot No"] = ngProcess
                    excelData["Process 3 Casing Packing"] = ngProcess
                    excelData["Process 3 Casing Packing Lot No"] = ngProcess
                    excelData["Process 3 M4x12 Screw"] = ngProcess
                    excelData["Process 3 M4x12 Screw Lot No"] = ngProcess
                    excelData["Process 3 Csb L"] = ngProcess
                    excelData["Process 3 Csb L Lot No"] = ngProcess
                    excelData["Process 3 Csb R"] = ngProcess
                    excelData["Process 3 Csb R Lot No"] = ngProcess
                    excelData["Process 3 Head Packing"] = ngProcess
                    excelData["Process 3 Head Packing Lot No"] = ngProcess
                    excelData["Process 3 ST"] = ngProcess
                    excelData["Process 3 Actual Time"] = ngProcess
                    excelData["Process 3 NG Cause"] = ngProcess
                    excelData["Process 3 Repaired Action"] = ngProcess

                    excelData["Process 6 Model Code"] = ngProcess
                    excelData["Process 6 S/N"] = ngProcess
                    excelData["Process 6 ID"] = ngProcess
                    excelData["Process 6 NAME"] = ngProcess
                    excelData["Process 6 Regular/Contractual"] = ngProcess
                    excelData["Process 6 Vinyl"] = ngProcess
                    excelData["Process 6 Vinyl Lot No"] = ngProcess
                    excelData["Process 6 ST"] = ngProcess
                    excelData["Process 6 Actual Time"] = ngProcess
                    excelData["Process 6 NG Cause"] = ngProcess
                    excelData["Process 6 Repaired Action"] = ngProcess
                    
                elif process6Status == "NG":
                    repairedProcess = "REPAIRED AT PROCESS4"
                    process4Row += 1
                    process5Row += 1
                    process6Row += 1
                    piRow += 1

                    excelData["Process 1 Model Code"] = repairedProcess
                    excelData["Process 1 S/N"] = repairedProcess
                    excelData["Process 1 ID"] = repairedProcess
                    excelData["Process 1 NAME"] = repairedProcess
                    excelData["Process 1 Regular/Contractual"] = repairedProcess
                    excelData["Process 1 Em2p"] = repairedProcess
                    excelData["Process 1 Em2p Lot No"] = repairedProcess
                    excelData["Process 1 Em3p"] = repairedProcess
                    excelData["Process 1 Em3p Lot No"] = repairedProcess
                    excelData["Process 1 Harness"] = repairedProcess
                    excelData["Process 1 Harness Lot No"] = repairedProcess
                    excelData["Process 1 Frame"] = repairedProcess
                    excelData["Process 1 Frame Lot No"] = repairedProcess 
                    excelData["Process 1 Bushing"] = repairedProcess
                    excelData["Process 1 Bushing Lot No"] = repairedProcess
                    excelData["Process 1 ST"] = repairedProcess
                    excelData["Process 1 Actual Time"] = repairedProcess
                    excelData["Process 1 NG Cause"] = repairedProcess
                    excelData["Process 1 Repaired Action"] = repairedProcess  

                    excelData["Process 2 Model Code"] = repairedProcess
                    excelData["Process 2 S/N"] = repairedProcess
                    excelData["Process 2 ID"] = repairedProcess
                    excelData["Process 2 NAME"] = repairedProcess
                    excelData["Process 2 Regular/Contractual"] = repairedProcess
                    excelData["Process 2 M4x40 Screw"] = repairedProcess
                    excelData["Process 2 M4x40 Screw Lot No"] = repairedProcess
                    excelData["Process 2 Rod Blk"] = repairedProcess
                    excelData["Process 2 Rod Blk Lot No"] = repairedProcess
                    excelData["Process 2 Df Blk"] = repairedProcess
                    excelData["Process 2 Df Blk Lot No"] = repairedProcess
                    excelData["Process 2 Df Ring"] = repairedProcess
                    excelData["Process 2 Df Ring Lot No"] = repairedProcess
                    excelData["Process 2 Washer"] = repairedProcess
                    excelData["Process 2 Washer Lot No"] = repairedProcess
                    excelData["Process 2 Lock Nut"] = repairedProcess
                    excelData["Process 2 Lock Nut Lot No"] = repairedProcess
                    excelData["Process 2 ST"] = repairedProcess
                    excelData["Process 2 Actual Time"] = repairedProcess
                    excelData["Process 2 NG Cause"] = repairedProcess
                    excelData["Process 2 Repaired Action"] = repairedProcess

                    excelData["Process 3 Model Code"] = repairedProcess
                    excelData["Process 3 S/N"] = repairedProcess
                    excelData["Process 3 ID"] = repairedProcess
                    excelData["Process 3 NAME"] = repairedProcess
                    excelData["Process 3 Regular/Contractual"] = repairedProcess
                    excelData["Process 3 Frame Gasket"] = repairedProcess
                    excelData["Process 3 Frame Gasket Lot No"] = repairedProcess
                    excelData["Process 3 Casing Block"] = repairedProcess
                    excelData["Process 3 Casing Block Lot No"] = repairedProcess
                    excelData["Process 3 Casing Block Inspection 1 Average Data"] = repairedProcess
                    excelData["Process 3 Casing Block Inspection 1 Minimum Data"] = repairedProcess
                    excelData["Process 3 Casing Block Inspection 1 Maximum Data"] = repairedProcess
                    excelData["Process 3 Casing Gasket"] = repairedProcess
                    excelData["Process 3 Casing Gasket Lot No"] = repairedProcess
                    excelData["Process 3 M4x16 Screw 1"] = repairedProcess
                    excelData["Process 3 M4x16 Screw 1 Lot No"] = repairedProcess
                    excelData["Process 3 M4x16 Screw 2"] = repairedProcess
                    excelData["Process 3 M4x16 Screw 2 Lot No"] = repairedProcess
                    excelData["Process 3 Ball Cushion"] = repairedProcess
                    excelData["Process 3 Ball Cushion Lot No"] = repairedProcess
                    excelData["Process 3 Frame Cover"] = repairedProcess
                    excelData["Process 3 Frame Cover Lot No"] = repairedProcess
                    excelData["Process 3 Partition Board"] = repairedProcess
                    excelData["Process 3 Partition Board Lot No"] = repairedProcess
                    excelData["Process 3 Built In Tube 1"] = repairedProcess
                    excelData["Process 3 Built In Tube 1 Lot No"] = repairedProcess
                    excelData["Process 3 Built In Tube 2"] = repairedProcess
                    excelData["Process 3 Built In Tube 2 Lot No"] = repairedProcess
                    excelData["Process 3 Head Cover"] = repairedProcess
                    excelData["Process 3 Head Cover Lot No"] = repairedProcess
                    excelData["Process 3 Casing Packing"] = repairedProcess
                    excelData["Process 3 Casing Packing Lot No"] = repairedProcess
                    excelData["Process 3 M4x12 Screw"] = repairedProcess
                    excelData["Process 3 M4x12 Screw Lot No"] = repairedProcess
                    excelData["Process 3 Csb L"] = repairedProcess
                    excelData["Process 3 Csb L Lot No"] = repairedProcess
                    excelData["Process 3 Csb R"] = repairedProcess
                    excelData["Process 3 Csb R Lot No"] = repairedProcess
                    excelData["Process 3 Head Packing"] = repairedProcess
                    excelData["Process 3 Head Packing Lot No"] = repairedProcess
                    excelData["Process 3 ST"] = repairedProcess
                    excelData["Process 3 Actual Time"] = repairedProcess
                    excelData["Process 3 NG Cause"] = repairedProcess
                    excelData["Process 3 Repaired Action"] = repairedProcess
                    
            elif process5Status == "Repaired":
                if process6Status == "NG":
                    repairedProcess = "REPAIRED AT PROCESS5"
                    process5Row += 1
                    process6Row += 1
                    piRow += 1

                    excelData["Process 1 Model Code"] = repairedProcess
                    excelData["Process 1 S/N"] = repairedProcess
                    excelData["Process 1 ID"] = repairedProcess
                    excelData["Process 1 NAME"] = repairedProcess
                    excelData["Process 1 Regular/Contractual"] = repairedProcess
                    excelData["Process 1 Em2p"] = repairedProcess
                    excelData["Process 1 Em2p Lot No"] = repairedProcess
                    excelData["Process 1 Em3p"] = repairedProcess
                    excelData["Process 1 Em3p Lot No"] = repairedProcess
                    excelData["Process 1 Harness"] = repairedProcess
                    excelData["Process 1 Harness Lot No"] = repairedProcess
                    excelData["Process 1 Frame"] = repairedProcess
                    excelData["Process 1 Frame Lot No"] = repairedProcess 
                    excelData["Process 1 Bushing"] = repairedProcess
                    excelData["Process 1 Bushing Lot No"] = repairedProcess
                    excelData["Process 1 ST"] = repairedProcess
                    excelData["Process 1 Actual Time"] = repairedProcess
                    excelData["Process 1 NG Cause"] = repairedProcess
                    excelData["Process 1 Repaired Action"] = repairedProcess  

                    excelData["Process 2 Model Code"] = repairedProcess
                    excelData["Process 2 S/N"] = repairedProcess
                    excelData["Process 2 ID"] = repairedProcess
                    excelData["Process 2 NAME"] = repairedProcess
                    excelData["Process 2 Regular/Contractual"] = repairedProcess
                    excelData["Process 2 M4x40 Screw"] = repairedProcess
                    excelData["Process 2 M4x40 Screw Lot No"] = repairedProcess
                    excelData["Process 2 Rod Blk"] = repairedProcess
                    excelData["Process 2 Rod Blk Lot No"] = repairedProcess
                    excelData["Process 2 Df Blk"] = repairedProcess
                    excelData["Process 2 Df Blk Lot No"] = repairedProcess
                    excelData["Process 2 Df Ring"] = repairedProcess
                    excelData["Process 2 Df Ring Lot No"] = repairedProcess
                    excelData["Process 2 Washer"] = repairedProcess
                    excelData["Process 2 Washer Lot No"] = repairedProcess
                    excelData["Process 2 Lock Nut"] = repairedProcess
                    excelData["Process 2 Lock Nut Lot No"] = repairedProcess
                    excelData["Process 2 ST"] = repairedProcess
                    excelData["Process 2 Actual Time"] = repairedProcess
                    excelData["Process 2 NG Cause"] = repairedProcess
                    excelData["Process 2 Repaired Action"] = repairedProcess

                    excelData["Process 3 Model Code"] = repairedProcess
                    excelData["Process 3 S/N"] = repairedProcess
                    excelData["Process 3 ID"] = repairedProcess
                    excelData["Process 3 NAME"] = repairedProcess
                    excelData["Process 3 Regular/Contractual"] = repairedProcess
                    excelData["Process 3 Frame Gasket"] = repairedProcess
                    excelData["Process 3 Frame Gasket Lot No"] = repairedProcess
                    excelData["Process 3 Casing Block"] = repairedProcess
                    excelData["Process 3 Casing Block Lot No"] = repairedProcess
                    excelData["Process 3 Casing Block Inspection 1 Average Data"] = repairedProcess
                    excelData["Process 3 Casing Block Inspection 1 Minimum Data"] = repairedProcess
                    excelData["Process 3 Casing Block Inspection 1 Maximum Data"] = repairedProcess
                    excelData["Process 3 Casing Gasket"] = repairedProcess
                    excelData["Process 3 Casing Gasket Lot No"] = repairedProcess
                    excelData["Process 3 M4x16 Screw 1"] = repairedProcess
                    excelData["Process 3 M4x16 Screw 1 Lot No"] = repairedProcess
                    excelData["Process 3 M4x16 Screw 2"] = repairedProcess
                    excelData["Process 3 M4x16 Screw 2 Lot No"] = repairedProcess
                    excelData["Process 3 Ball Cushion"] = repairedProcess
                    excelData["Process 3 Ball Cushion Lot No"] = repairedProcess
                    excelData["Process 3 Frame Cover"] = repairedProcess
                    excelData["Process 3 Frame Cover Lot No"] = repairedProcess
                    excelData["Process 3 Partition Board"] = repairedProcess
                    excelData["Process 3 Partition Board Lot No"] = repairedProcess
                    excelData["Process 3 Built In Tube 1"] = repairedProcess
                    excelData["Process 3 Built In Tube 1 Lot No"] = repairedProcess
                    excelData["Process 3 Built In Tube 2"] = repairedProcess
                    excelData["Process 3 Built In Tube 2 Lot No"] = repairedProcess
                    excelData["Process 3 Head Cover"] = repairedProcess
                    excelData["Process 3 Head Cover Lot No"] = repairedProcess
                    excelData["Process 3 Casing Packing"] = repairedProcess
                    excelData["Process 3 Casing Packing Lot No"] = repairedProcess
                    excelData["Process 3 M4x12 Screw"] = repairedProcess
                    excelData["Process 3 M4x12 Screw Lot No"] = repairedProcess
                    excelData["Process 3 Csb L"] = repairedProcess
                    excelData["Process 3 Csb L Lot No"] = repairedProcess
                    excelData["Process 3 Csb R"] = repairedProcess
                    excelData["Process 3 Csb R Lot No"] = repairedProcess
                    excelData["Process 3 Head Packing"] = repairedProcess
                    excelData["Process 3 Head Packing Lot No"] = repairedProcess
                    excelData["Process 3 ST"] = repairedProcess
                    excelData["Process 3 Actual Time"] = repairedProcess
                    excelData["Process 3 NG Cause"] = repairedProcess
                    excelData["Process 3 Repaired Action"] = repairedProcess

                    excelData["Process 4 Model Code"] = repairedProcess
                    excelData["Process 4 S/N"] = repairedProcess
                    excelData["Process 4 ID"] = repairedProcess
                    excelData["Process 4 NAME"] = repairedProcess
                    excelData["Process 4 Regular/Contractual"] = repairedProcess
                    excelData["Process 4 Tank"] = repairedProcess
                    excelData["Process 4 Tank Lot No"] = repairedProcess
                    excelData["Process 4 Upper Housing"] = repairedProcess
                    excelData["Process 4 Upper Housing Lot No"] = repairedProcess
                    excelData["Process 4 Cord Hook" ] = repairedProcess
                    excelData["Process 4 Cord Hook Lot No"] = repairedProcess
                    excelData["Process 4 M4x16 Screw"] = repairedProcess
                    excelData["Process 4 M4x16 Screw Lot No"] = repairedProcess
                    excelData["Process 4 Tank Gasket"] = repairedProcess
                    excelData["Process 4 Tank Gasket Lot No"] = repairedProcess
                    excelData["Process 4 Tank Cover"] = repairedProcess
                    excelData["Process 4 Tank Cover Lot No"] = repairedProcess
                    excelData["Process 4 Housing Gasket"] = repairedProcess
                    excelData["Process 4 Housing Gasket Lot No"] = repairedProcess
                    excelData["Process 4 M4x40 Screw"] = repairedProcess
                    excelData["Process 4 M4x40 Screw Lot No"] = repairedProcess
                    excelData["Process 4 PartitionGasket"] = repairedProcess
                    excelData["Process 4 PartitionGasket Lot No"] = repairedProcess
                    excelData["Process 4 M4x12 Screw"] = repairedProcess
                    excelData["Process 4 M4x12 Screw Lot No"] = repairedProcess
                    excelData["Process 4 Muffler"] = repairedProcess
                    excelData["Process 4 Muffler Lot No"] = repairedProcess
                    excelData["Process 4 Muffler Gasket"] = repairedProcess
                    excelData["Process 4 Muffler Gasket Lot No"] = repairedProcess
                    excelData["Process 4 VCR"] = repairedProcess
                    excelData["Process 4 VCR Lot No"] = repairedProcess
                    excelData["Process 4 ST"] = repairedProcess
                    excelData["Process 4 Actual Time"] = repairedProcess
                    excelData["Process 4 NG Cause"] = repairedProcess
                    excelData["Process 4 Repaired Action"] = repairedProcess
                
            # elif process6Status == "Repaired":
            #     pass

            process1Status = ""
            process2Status = ""
            process3Status = ""
            process4Status = ""
            process5Status = ""
            process6Status = ""

        if process1Status == "NG":
            ngProcess = "NG AT PROCESS1"
            process1Row += 1

            excelData["DATE"] = ngProcess
            excelData["TIME"] = ngProcess
            excelData["MODEL CODE"] = ngProcess
            excelData["PROCESS S/N"] = ngProcess
            excelData["S/N"] = ngProcess
            excelData["PASS/NG"] = ngProcess
            excelData["VOLTAGE MAX (V)"] = ngProcess
            excelData["WATTAGE MAX (W)"] = ngProcess
            excelData["CLOSED PRESSURE_MAX (kPa)"] = ngProcess
            excelData["VOLTAGE Middle (V)"] = ngProcess
            excelData["WATTAGE Middle (W)"] = ngProcess
            excelData["AMPERAGE Middle (A)"] = ngProcess
            excelData["CLOSED PRESSURE Middle (kPa)"] = ngProcess
            excelData["dB(A) 1"] = ngProcess
            excelData["dB(A) 2"] = ngProcess
            excelData["dB(A) 3"] = ngProcess
            excelData["VOLTAGE MIN (V)"] = ngProcess
            excelData["WATTAGE MIN (W)"] = ngProcess
            excelData["CLOSED PRESSURE MIN (kPa)"] = ngProcess

            excelData["Process 2 Model Code"] = ngProcess
            excelData["Process 2 S/N"] = ngProcess
            excelData["Process 2 ID"] = ngProcess
            excelData["Process 2 NAME"] = ngProcess
            excelData["Process 2 Regular/Contractual"] = ngProcess
            excelData["Process 2 M4x40 Screw"] = ngProcess
            excelData["Process 2 M4x40 Screw Lot No"] = ngProcess
            excelData["Process 2 Rod Blk"] = ngProcess
            excelData["Process 2 Rod Blk Lot No"] = ngProcess
            excelData["Process 2 Df Blk"] = ngProcess
            excelData["Process 2 Df Blk Lot No"] = ngProcess
            excelData["Process 2 Df Ring"] = ngProcess
            excelData["Process 2 Df Ring Lot No"] = ngProcess
            excelData["Process 2 Washer"] = ngProcess
            excelData["Process 2 Washer Lot No"] = ngProcess
            excelData["Process 2 Lock Nut"] = ngProcess
            excelData["Process 2 Lock Nut Lot No"] = ngProcess
            excelData["Process 2 ST"] = ngProcess
            excelData["Process 2 Actual Time"] = ngProcess
            excelData["Process 2 NG Cause"] = ngProcess
            excelData["Process 2 Repaired Action"] = ngProcess

            excelData["Process 3 Model Code"] = ngProcess
            excelData["Process 3 S/N"] = ngProcess
            excelData["Process 3 ID"] = ngProcess
            excelData["Process 3 NAME"] = ngProcess
            excelData["Process 3 Regular/Contractual"] = ngProcess
            excelData["Process 3 Frame Gasket"] = ngProcess
            excelData["Process 3 Frame Gasket Lot No"] = ngProcess
            excelData["Process 3 Casing Block"] = ngProcess
            excelData["Process 3 Casing Block Lot No"] = ngProcess
            excelData["Process 3 Casing Block Inspection 1 Average Data"] = ngProcess
            excelData["Process 3 Casing Block Inspection 1 Minimum Data"] = ngProcess
            excelData["Process 3 Casing Block Inspection 1 Maximum Data"] = ngProcess
            excelData["Process 3 Casing Gasket"] = ngProcess
            excelData["Process 3 Casing Gasket Lot No"] = ngProcess
            excelData["Process 3 M4x16 Screw 1"] = ngProcess
            excelData["Process 3 M4x16 Screw 1 Lot No"] = ngProcess
            excelData["Process 3 M4x16 Screw 2"] = ngProcess
            excelData["Process 3 M4x16 Screw 2 Lot No"] = ngProcess
            excelData["Process 3 Ball Cushion"] = ngProcess
            excelData["Process 3 Ball Cushion Lot No"] = ngProcess
            excelData["Process 3 Frame Cover"] = ngProcess
            excelData["Process 3 Frame Cover Lot No"] = ngProcess
            excelData["Process 3 Partition Board"] = ngProcess
            excelData["Process 3 Partition Board Lot No"] = ngProcess
            excelData["Process 3 Built In Tube 1"] = ngProcess
            excelData["Process 3 Built In Tube 1 Lot No"] = ngProcess
            excelData["Process 3 Built In Tube 2"] = ngProcess
            excelData["Process 3 Built In Tube 2 Lot No"] = ngProcess
            excelData["Process 3 Head Cover"] = ngProcess
            excelData["Process 3 Head Cover Lot No"] = ngProcess
            excelData["Process 3 Casing Packing"] = ngProcess
            excelData["Process 3 Casing Packing Lot No"] = ngProcess
            excelData["Process 3 M4x12 Screw"] = ngProcess
            excelData["Process 3 M4x12 Screw Lot No"] = ngProcess
            excelData["Process 3 Csb L"] = ngProcess
            excelData["Process 3 Csb L Lot No"] = ngProcess
            excelData["Process 3 Csb R"] = ngProcess
            excelData["Process 3 Csb R Lot No"] = ngProcess
            excelData["Process 3 Head Packing"] = ngProcess
            excelData["Process 3 Head Packing Lot No"] = ngProcess
            excelData["Process 3 ST"] = ngProcess
            excelData["Process 3 Actual Time"] = ngProcess
            excelData["Process 3 NG Cause"] = ngProcess
            excelData["Process 3 Repaired Action"] = ngProcess

            excelData["Process 4 Model Code"] = ngProcess
            excelData["Process 4 S/N"] = ngProcess
            excelData["Process 4 ID"] = ngProcess
            excelData["Process 4 NAME"] = ngProcess
            excelData["Process 4 Regular/Contractual"] = ngProcess
            excelData["Process 4 Tank"] = ngProcess
            excelData["Process 4 Tank Lot No"] = ngProcess
            excelData["Process 4 Upper Housing"] = ngProcess
            excelData["Process 4 Upper Housing Lot No"] = ngProcess
            excelData["Process 4 Cord Hook" ] = ngProcess
            excelData["Process 4 Cord Hook Lot No"] = ngProcess
            excelData["Process 4 M4x16 Screw"] = ngProcess
            excelData["Process 4 M4x16 Screw Lot No"] = ngProcess
            excelData["Process 4 Tank Gasket"] = ngProcess
            excelData["Process 4 Tank Gasket Lot No"] = ngProcess
            excelData["Process 4 Tank Cover"] = ngProcess
            excelData["Process 4 Tank Cover Lot No"] = ngProcess
            excelData["Process 4 Housing Gasket"] = ngProcess
            excelData["Process 4 Housing Gasket Lot No"] = ngProcess
            excelData["Process 4 M4x40 Screw"] = ngProcess
            excelData["Process 4 M4x40 Screw Lot No"] = ngProcess
            excelData["Process 4 PartitionGasket"] = ngProcess
            excelData["Process 4 PartitionGasket Lot No"] = ngProcess
            excelData["Process 4 M4x12 Screw"] = ngProcess
            excelData["Process 4 M4x12 Screw Lot No"] = ngProcess
            excelData["Process 4 Muffler"] = ngProcess
            excelData["Process 4 Muffler Lot No"] = ngProcess
            excelData["Process 4 Muffler Gasket"] = ngProcess
            excelData["Process 4 Muffler Gasket Lot No"] = ngProcess
            excelData["Process 4 VCR"] = ngProcess
            excelData["Process 4 VCR Lot No"] = ngProcess
            excelData["Process 4 ST"] = ngProcess
            excelData["Process 4 Actual Time"] = ngProcess
            excelData["Process 4 NG Cause"] = ngProcess
            excelData["Process 4 Repaired Action"] = ngProcess

            excelData["Process 5 Model Code"] = ngProcess
            excelData["Process 5 S/N"] = ngProcess
            excelData["Process 5 ID"] = ngProcess
            excelData["Process 5 NAME"] = ngProcess
            excelData["Process 5 Regular/Contractual"] = ngProcess
            excelData["Process 5 Rating Label"] = ngProcess
            excelData["Process 5 Rating Label Lot No"] = ngProcess
            excelData["Process 5 ST"] = ngProcess
            excelData["Process 5 Actual Time"] = ngProcess
            excelData["Process 5 NG Cause"] = ngProcess
            excelData["Process 5 Repaired Action"] = ngProcess

            excelData["Process 6 Model Code"] = ngProcess
            excelData["Process 6 S/N"] = ngProcess
            excelData["Process 6 ID"] = ngProcess
            excelData["Process 6 NAME"] = ngProcess
            excelData["Process 6 Regular/Contractual"] = ngProcess
            excelData["Process 6 Vinyl"] = ngProcess
            excelData["Process 6 Vinyl Lot No"] = ngProcess
            excelData["Process 6 ST"] = ngProcess
            excelData["Process 6 Actual Time"] = ngProcess
            excelData["Process 6 NG Cause"] = ngProcess
            excelData["Process 6 Repaired Action"] = ngProcess 
            
        if process2Status == "NG":
            print("ng")
            ngProcess = "NG AT PROCESS2"
            process2Row += 1
            excelData["DATE"] = ngProcess
            excelData["TIME"] = ngProcess
            excelData["MODEL CODE"] = ngProcess
            excelData["PROCESS S/N"] = ngProcess
            excelData["S/N"] = ngProcess
            excelData["PASS/NG"] = ngProcess
            excelData["VOLTAGE MAX (V)"] = ngProcess
            excelData["WATTAGE MAX (W)"] = ngProcess
            excelData["CLOSED PRESSURE_MAX (kPa)"] = ngProcess
            excelData["VOLTAGE Middle (V)"] = ngProcess
            excelData["WATTAGE Middle (W)"] = ngProcess
            excelData["AMPERAGE Middle (A)"] = ngProcess
            excelData["CLOSED PRESSURE Middle (kPa)"] = ngProcess
            excelData["dB(A) 1"] = ngProcess
            excelData["dB(A) 2"] = ngProcess
            excelData["dB(A) 3"] = ngProcess
            excelData["VOLTAGE MIN (V)"] = ngProcess
            excelData["WATTAGE MIN (W)"] = ngProcess
            excelData["CLOSED PRESSURE MIN (kPa)"] = ngProcess
    
            excelData["Process 3 Model Code"] = ngProcess
            excelData["Process 3 S/N"] = ngProcess
            excelData["Process 3 ID"] = ngProcess
            excelData["Process 3 NAME"] = ngProcess
            excelData["Process 3 Regular/Contractual"] = ngProcess
            excelData["Process 3 Frame Gasket"] = ngProcess
            excelData["Process 3 Frame Gasket Lot No"] = ngProcess
            excelData["Process 3 Casing Block"] = ngProcess
            excelData["Process 3 Casing Block Lot No"] = ngProcess
            excelData["Process 3 Casing Block Inspection 1 Average Data"] = ngProcess
            excelData["Process 3 Casing Block Inspection 1 Minimum Data"] = ngProcess
            excelData["Process 3 Casing Block Inspection 1 Maximum Data"] = ngProcess
            excelData["Process 3 Casing Gasket"] = ngProcess
            excelData["Process 3 Casing Gasket Lot No"] = ngProcess
            excelData["Process 3 M4x16 Screw 1"] = ngProcess
            excelData["Process 3 M4x16 Screw 1 Lot No"] = ngProcess
            excelData["Process 3 M4x16 Screw 2"] = ngProcess
            excelData["Process 3 M4x16 Screw 2 Lot No"] = ngProcess
            excelData["Process 3 Ball Cushion"] = ngProcess
            excelData["Process 3 Ball Cushion Lot No"] = ngProcess
            excelData["Process 3 Frame Cover"] = ngProcess
            excelData["Process 3 Frame Cover Lot No"] = ngProcess
            excelData["Process 3 Partition Board"] = ngProcess
            excelData["Process 3 Partition Board Lot No"] = ngProcess
            excelData["Process 3 Built In Tube 1"] = ngProcess
            excelData["Process 3 Built In Tube 1 Lot No"] = ngProcess
            excelData["Process 3 Built In Tube 2"] = ngProcess
            excelData["Process 3 Built In Tube 2 Lot No"] = ngProcess
            excelData["Process 3 Head Cover"] = ngProcess
            excelData["Process 3 Head Cover Lot No"] = ngProcess
            excelData["Process 3 Casing Packing"] = ngProcess
            excelData["Process 3 Casing Packing Lot No"] = ngProcess
            excelData["Process 3 M4x12 Screw"] = ngProcess
            excelData["Process 3 M4x12 Screw Lot No"] = ngProcess
            excelData["Process 3 Csb L"] = ngProcess
            excelData["Process 3 Csb L Lot No"] = ngProcess
            excelData["Process 3 Csb R"] = ngProcess
            excelData["Process 3 Csb R Lot No"] = ngProcess
            excelData["Process 3 Head Packing"] = ngProcess
            excelData["Process 3 Head Packing Lot No"] = ngProcess
            excelData["Process 3 ST"] = ngProcess
            excelData["Process 3 Actual Time"] = ngProcess
            excelData["Process 3 NG Cause"] = ngProcess
            excelData["Process 3 Repaired Action"] = ngProcess

            excelData["Process 4 Model Code"] = ngProcess
            excelData["Process 4 S/N"] = ngProcess
            excelData["Process 4 ID"] = ngProcess
            excelData["Process 4 NAME"] = ngProcess
            excelData["Process 4 Regular/Contractual"] = ngProcess
            excelData["Process 4 Tank"] = ngProcess
            excelData["Process 4 Tank Lot No"] = ngProcess
            excelData["Process 4 Upper Housing"] = ngProcess
            excelData["Process 4 Upper Housing Lot No"] = ngProcess
            excelData["Process 4 Cord Hook" ] = ngProcess
            excelData["Process 4 Cord Hook Lot No"] = ngProcess
            excelData["Process 4 M4x16 Screw"] = ngProcess
            excelData["Process 4 M4x16 Screw Lot No"] = ngProcess
            excelData["Process 4 Tank Gasket"] = ngProcess
            excelData["Process 4 Tank Gasket Lot No"] = ngProcess
            excelData["Process 4 Tank Cover"] = ngProcess
            excelData["Process 4 Tank Cover Lot No"] = ngProcess
            excelData["Process 4 Housing Gasket"] = ngProcess
            excelData["Process 4 Housing Gasket Lot No"] = ngProcess
            excelData["Process 4 M4x40 Screw"] = ngProcess
            excelData["Process 4 M4x40 Screw Lot No"] = ngProcess
            excelData["Process 4 PartitionGasket"] = ngProcess
            excelData["Process 4 PartitionGasket Lot No"] = ngProcess
            excelData["Process 4 M4x12 Screw"] = ngProcess
            excelData["Process 4 M4x12 Screw Lot No"] = ngProcess
            excelData["Process 4 Muffler"] = ngProcess
            excelData["Process 4 Muffler Lot No"] = ngProcess
            excelData["Process 4 Muffler Gasket"] = ngProcess
            excelData["Process 4 Muffler Gasket Lot No"] = ngProcess
            excelData["Process 4 VCR"] = ngProcess
            excelData["Process 4 VCR Lot No"] = ngProcess
            excelData["Process 4 ST"] = ngProcess
            excelData["Process 4 Actual Time"] = ngProcess
            excelData["Process 4 NG Cause"] = ngProcess
            excelData["Process 4 Repaired Action"] = ngProcess

            excelData["Process 5 Model Code"] = ngProcess
            excelData["Process 5 S/N"] = ngProcess
            excelData["Process 5 ID"] = ngProcess
            excelData["Process 5 NAME"] = ngProcess
            excelData["Process 5 Regular/Contractual"] = ngProcess
            excelData["Process 5 Rating Label"] = ngProcess
            excelData["Process 5 Rating Label Lot No"] = ngProcess
            excelData["Process 5 ST"] = ngProcess
            excelData["Process 5 Actual Time"] = ngProcess
            excelData["Process 5 NG Cause"] = ngProcess
            excelData["Process 5 Repaired Action"] = ngProcess

            excelData["Process 6 Model Code"] = ngProcess
            excelData["Process 6 S/N"] = ngProcess
            excelData["Process 6 ID"] = ngProcess
            excelData["Process 6 NAME"] = ngProcess
            excelData["Process 6 Regular/Contractual"] = ngProcess
            excelData["Process 6 Vinyl"] = ngProcess
            excelData["Process 6 Vinyl Lot No"] = ngProcess
            excelData["Process 6 ST"] = ngProcess
            excelData["Process 6 Actual Time"] = ngProcess
            excelData["Process 6 NG Cause"] = ngProcess
            excelData["Process 6 Repaired Action"] = ngProcess

        if process3Status == "NG":
            ngProcess = "NG AT PROCESS3"
            process3Row += 1
            excelData["DATE"] = ngProcess
            excelData["TIME"] = ngProcess
            excelData["MODEL CODE"] = ngProcess
            excelData["PROCESS S/N"] = ngProcess
            excelData["S/N"] = ngProcess
            excelData["PASS/NG"] = ngProcess
            excelData["VOLTAGE MAX (V)"] = ngProcess
            excelData["WATTAGE MAX (W)"] = ngProcess
            excelData["CLOSED PRESSURE_MAX (kPa)"] = ngProcess
            excelData["VOLTAGE Middle (V)"] = ngProcess
            excelData["WATTAGE Middle (W)"] = ngProcess
            excelData["AMPERAGE Middle (A)"] = ngProcess
            excelData["CLOSED PRESSURE Middle (kPa)"] = ngProcess
            excelData["dB(A) 1"] = ngProcess
            excelData["dB(A) 2"] = ngProcess
            excelData["dB(A) 3"] = ngProcess
            excelData["VOLTAGE MIN (V)"] = ngProcess
            excelData["WATTAGE MIN (W)"] = ngProcess
            excelData["CLOSED PRESSURE MIN (kPa)"] = ngProcess

            excelData["Process 4 Model Code"] = ngProcess
            excelData["Process 4 S/N"] = ngProcess
            excelData["Process 4 ID"] = ngProcess
            excelData["Process 4 NAME"] = ngProcess
            excelData["Process 4 Regular/Contractual"] = ngProcess
            excelData["Process 4 Tank"] = ngProcess
            excelData["Process 4 Tank Lot No"] = ngProcess
            excelData["Process 4 Upper Housing"] = ngProcess
            excelData["Process 4 Upper Housing Lot No"] = ngProcess
            excelData["Process 4 Cord Hook" ] = ngProcess
            excelData["Process 4 Cord Hook Lot No"] = ngProcess
            excelData["Process 4 M4x16 Screw"] = ngProcess
            excelData["Process 4 M4x16 Screw Lot No"] = ngProcess
            excelData["Process 4 Tank Gasket"] = ngProcess
            excelData["Process 4 Tank Gasket Lot No"] = ngProcess
            excelData["Process 4 Tank Cover"] = ngProcess
            excelData["Process 4 Tank Cover Lot No"] = ngProcess
            excelData["Process 4 Housing Gasket"] = ngProcess
            excelData["Process 4 Housing Gasket Lot No"] = ngProcess
            excelData["Process 4 M4x40 Screw"] = ngProcess
            excelData["Process 4 M4x40 Screw Lot No"] = ngProcess
            excelData["Process 4 PartitionGasket"] = ngProcess
            excelData["Process 4 PartitionGasket Lot No"] = ngProcess
            excelData["Process 4 M4x12 Screw"] = ngProcess
            excelData["Process 4 M4x12 Screw Lot No"] = ngProcess
            excelData["Process 4 Muffler"] = ngProcess
            excelData["Process 4 Muffler Lot No"] = ngProcess
            excelData["Process 4 Muffler Gasket"] = ngProcess
            excelData["Process 4 Muffler Gasket Lot No"] = ngProcess
            excelData["Process 4 VCR"] = ngProcess
            excelData["Process 4 VCR Lot No"] = ngProcess
            excelData["Process 4 ST"] = ngProcess
            excelData["Process 4 Actual Time"] = ngProcess
            excelData["Process 4 NG Cause"] = ngProcess
            excelData["Process 4 Repaired Action"] = ngProcess

            excelData["Process 5 Model Code"] = ngProcess
            excelData["Process 5 S/N"] = ngProcess
            excelData["Process 5 ID"] = ngProcess
            excelData["Process 5 NAME"] = ngProcess
            excelData["Process 5 Regular/Contractual"] = ngProcess
            excelData["Process 5 Rating Label"] = ngProcess
            excelData["Process 5 Rating Label Lot No"] = ngProcess
            excelData["Process 5 ST"] = ngProcess
            excelData["Process 5 Actual Time"] = ngProcess
            excelData["Process 5 NG Cause"] = ngProcess
            excelData["Process 5 Repaired Action"] = ngProcess

            excelData["Process 6 Model Code"] = ngProcess
            excelData["Process 6 S/N"] = ngProcess
            excelData["Process 6 ID"] = ngProcess
            excelData["Process 6 NAME"] = ngProcess
            excelData["Process 6 Regular/Contractual"] = ngProcess
            excelData["Process 6 Vinyl"] = ngProcess
            excelData["Process 6 Vinyl Lot No"] = ngProcess
            excelData["Process 6 ST"] = ngProcess
            excelData["Process 6 Actual Time"] = ngProcess
            excelData["Process 6 NG Cause"] = ngProcess
            excelData["Process 6 Repaired Action"] = ngProcess

        if process4Status == "NG":
            ngProcess = "NG AT PROCESS4"
            process4Row += 1
            excelData["DATE"] = ngProcess
            excelData["TIME"] = ngProcess
            excelData["MODEL CODE"] = ngProcess
            excelData["PROCESS S/N"] = ngProcess
            excelData["S/N"] = ngProcess
            excelData["PASS/NG"] = ngProcess
            excelData["VOLTAGE MAX (V)"] = ngProcess
            excelData["WATTAGE MAX (W)"] = ngProcess
            excelData["CLOSED PRESSURE_MAX (kPa)"] = ngProcess
            excelData["VOLTAGE Middle (V)"] = ngProcess
            excelData["WATTAGE Middle (W)"] = ngProcess
            excelData["AMPERAGE Middle (A)"] = ngProcess
            excelData["CLOSED PRESSURE Middle (kPa)"] = ngProcess
            excelData["dB(A) 1"] = ngProcess
            excelData["dB(A) 2"] = ngProcess
            excelData["dB(A) 3"] = ngProcess
            excelData["VOLTAGE MIN (V)"] = ngProcess
            excelData["WATTAGE MIN (W)"] = ngProcess
            excelData["CLOSED PRESSURE MIN (kPa)"] = ngProcess

            excelData["Process 5 Model Code"] = ngProcess
            excelData["Process 5 S/N"] = ngProcess
            excelData["Process 5 ID"] = ngProcess
            excelData["Process 5 NAME"] = ngProcess
            excelData["Process 5 Regular/Contractual"] = ngProcess
            excelData["Process 5 Rating Label"] = ngProcess
            excelData["Process 5 Rating Label Lot No"] = ngProcess
            excelData["Process 5 ST"] = ngProcess
            excelData["Process 5 Actual Time"] = ngProcess
            excelData["Process 5 NG Cause"] = ngProcess
            excelData["Process 5 Repaired Action"] = ngProcess

            excelData["Process 6 Model Code"] = ngProcess
            excelData["Process 6 S/N"] = ngProcess
            excelData["Process 6 ID"] = ngProcess
            excelData["Process 6 NAME"] = ngProcess
            excelData["Process 6 Regular/Contractual"] = ngProcess
            excelData["Process 6 Vinyl"] = ngProcess
            excelData["Process 6 Vinyl Lot No"] = ngProcess
            excelData["Process 6 ST"] = ngProcess
            excelData["Process 6 Actual Time"] = ngProcess
            excelData["Process 6 NG Cause"] = ngProcess
            excelData["Process 6 Repaired Action"] = ngProcess

        if process5Status == "NG PRESSURE":
            ngProcess = "NG PRESSURE AT PROCESS5"
            process5Row += 1
            piRow += 1

            excelData["DATE"] = ngProcess
            excelData["TIME"] = ngProcess
            excelData["MODEL CODE"] = ngProcess
            excelData["PROCESS S/N"] = ngProcess
            excelData["S/N"] = ngProcess
            excelData["PASS/NG"] = ngProcess
            excelData["VOLTAGE MAX (V)"] = ngProcess
            excelData["WATTAGE MAX (W)"] = ngProcess
            excelData["CLOSED PRESSURE_MAX (kPa)"] = ngProcess
            excelData["VOLTAGE Middle (V)"] = ngProcess
            excelData["WATTAGE Middle (W)"] = ngProcess
            excelData["AMPERAGE Middle (A)"] = ngProcess
            excelData["CLOSED PRESSURE Middle (kPa)"] = ngProcess
            excelData["dB(A) 1"] = ngProcess
            excelData["dB(A) 2"] = ngProcess
            excelData["dB(A) 3"] = ngProcess
            excelData["VOLTAGE MIN (V)"] = ngProcess
            excelData["WATTAGE MIN (W)"] = ngProcess
            excelData["CLOSED PRESSURE MIN (kPa)"] = ngProcess

            excelData["Process 6 Model Code"] = ngProcess
            excelData["Process 6 S/N"] = ngProcess
            excelData["Process 6 ID"] = ngProcess
            excelData["Process 6 NAME"] = ngProcess
            excelData["Process 6 Regular/Contractual"] = ngProcess
            excelData["Process 6 Vinyl"] = ngProcess
            excelData["Process 6 Vinyl Lot No"] = ngProcess
            excelData["Process 6 ST"] = ngProcess
            excelData["Process 6 Actual Time"] = ngProcess
            excelData["Process 6 NG Cause"] = ngProcess
            excelData["Process 6 Repaired Action"] = ngProcess

        if process5Status == "NG":
            ngProcess = "NG AT PROCESS5"
            process5Row += 1
            piRow += 1

            # excelData["DATE"] = ngProcess
            # excelData["TIME"] = ngProcess
            # excelData["MODEL CODE"] = ngProcess
            # excelData["PROCESS S/N"] = ngProcess
            # excelData["S/N"] = ngProcess
            # excelData["PASS/NG"] = ngProcess
            # excelData["VOLTAGE MAX (V)"] = ngProcess
            # excelData["WATTAGE MAX (W)"] = ngProcess
            # excelData["CLOSED PRESSURE_MAX (kPa)"] = ngProcess
            # excelData["VOLTAGE Middle (V)"] = ngProcess
            # excelData["WATTAGE Middle (W)"] = ngProcess
            # excelData["AMPERAGE Middle (A)"] = ngProcess
            # excelData["CLOSED PRESSURE Middle (kPa)"] = ngProcess
            # excelData["dB(A) 1"] = ngProcess
            # excelData["dB(A) 2"] = ngProcess
            # excelData["dB(A) 3"] = ngProcess
            # excelData["VOLTAGE MIN (V)"] = ngProcess
            # excelData["WATTAGE MIN (W)"] = ngProcess
            # excelData["CLOSED PRESSURE MIN (kPa)"] = ngProcess

            excelData["Process 6 Model Code"] = ngProcess
            excelData["Process 6 S/N"] = ngProcess
            excelData["Process 6 ID"] = ngProcess
            excelData["Process 6 NAME"] = ngProcess
            excelData["Process 6 Regular/Contractual"] = ngProcess
            excelData["Process 6 Vinyl"] = ngProcess
            excelData["Process 6 Vinyl Lot No"] = ngProcess
            excelData["Process 6 ST"] = ngProcess
            excelData["Process 6 Actual Time"] = ngProcess
            excelData["Process 6 NG Cause"] = ngProcess
            excelData["Process 6 Repaired Action"] = ngProcess

        if process6Status == "NG":
            ngProcess = "NG AT PROCESS6"
            process6Row += 1
            excelData["DATE"] = ngProcess
            excelData["TIME"] = ngProcess
            excelData["MODEL CODE"] = ngProcess
            excelData["PROCESS S/N"] = ngProcess
            excelData["S/N"] = ngProcess
            excelData["PASS/NG"] = ngProcess
            excelData["VOLTAGE MAX (V)"] = ngProcess
            excelData["WATTAGE MAX (W)"] = ngProcess
            excelData["CLOSED PRESSURE_MAX (kPa)"] = ngProcess
            excelData["VOLTAGE Middle (V)"] = ngProcess
            excelData["WATTAGE Middle (W)"] = ngProcess
            excelData["AMPERAGE Middle (A)"] = ngProcess
            excelData["CLOSED PRESSURE Middle (kPa)"] = ngProcess
            excelData["dB(A) 1"] = ngProcess
            excelData["dB(A) 2"] = ngProcess
            excelData["dB(A) 3"] = ngProcess
            excelData["VOLTAGE MIN (V)"] = ngProcess
            excelData["WATTAGE MIN (W)"] = ngProcess
            excelData["CLOSED PRESSURE MIN (kPa)"] = ngProcess

        if process1Status == "Repaired":
            repairedProcess = "REPAIRED AT PROCESS1"
            process1Row += 1
            process2Row += 1
            process3Row += 1
            process4Row += 1
            process5Row += 1
            process6Row += 1
            piRow += 1

            # excelData["DATE"] = repairedProcess
            # excelData["TIME"] = repairedProcess
            # excelData["MODEL CODE"] = repairedProcess
            # excelData["PROCESS S/N"] = repairedProcess
            # excelData["S/N"] = repairedProcess
            # excelData["PASS/NG"] = repairedProcess
            # excelData["VOLTAGE MAX (V)"] = repairedProcess
            # excelData["WATTAGE MAX (W)"] = repairedProcess
            # excelData["CLOSED PRESSURE_MAX (kPa)"] = repairedProcess
            # excelData["VOLTAGE Middle (V)"] = repairedProcess
            # excelData["WATTAGE Middle (W)"] = repairedProcess
            # excelData["AMPERAGE Middle (A)"] = repairedProcess
            # excelData["CLOSED PRESSURE Middle (kPa)"] = repairedProcess
            # excelData["dB(A) 1"] = repairedProcess
            # excelData["dB(A) 2"] = repairedProcess
            # excelData["dB(A) 3"] = repairedProcess
            # excelData["VOLTAGE MIN (V)"] = repairedProcess
            # excelData["WATTAGE MIN (W)"] = repairedProcess
            # excelData["CLOSED PRESSURE MIN (kPa)"] = repairedProcess

        if process2Status == "Repaired":
            repairedProcess = "REPAIRED AT PROCESS2"
            process2Row += 1
            process3Row += 1
            process4Row += 1
            process5Row += 1
            process6Row += 1
            piRow += 1

            # excelData["DATE"] = repairedProcess
            # excelData["TIME"] = repairedProcess
            # excelData["MODEL CODE"] = repairedProcess
            # excelData["PROCESS S/N"] = repairedProcess
            # excelData["S/N"] = repairedProcess
            # excelData["PASS/NG"] = repairedProcess
            # excelData["VOLTAGE MAX (V)"] = repairedProcess
            # excelData["WATTAGE MAX (W)"] = repairedProcess
            # excelData["CLOSED PRESSURE_MAX (kPa)"] = repairedProcess
            # excelData["VOLTAGE Middle (V)"] = repairedProcess
            # excelData["WATTAGE Middle (W)"] = repairedProcess
            # excelData["AMPERAGE Middle (A)"] = repairedProcess
            # excelData["CLOSED PRESSURE Middle (kPa)"] = repairedProcess
            # excelData["dB(A) 1"] = repairedProcess
            # excelData["dB(A) 2"] = repairedProcess
            # excelData["dB(A) 3"] = repairedProcess
            # excelData["VOLTAGE MIN (V)"] = repairedProcess
            # excelData["WATTAGE MIN (W)"] = repairedProcess
            # excelData["CLOSED PRESSURE MIN (kPa)"] = repairedProcess

            excelData["Process 1 Model Code"] = repairedProcess
            excelData["Process 1 S/N"] = repairedProcess
            excelData["Process 1 ID"] = repairedProcess
            excelData["Process 1 NAME"] = repairedProcess
            excelData["Process 1 Regular/Contractual"] = repairedProcess
            excelData["Process 1 Em2p"] = repairedProcess
            excelData["Process 1 Em2p Lot No"] = repairedProcess
            excelData["Process 1 Em3p"] = repairedProcess
            excelData["Process 1 Em3p Lot No"] = repairedProcess
            excelData["Process 1 Harness"] = repairedProcess
            excelData["Process 1 Harness Lot No"] = repairedProcess
            excelData["Process 1 Frame"] = repairedProcess
            excelData["Process 1 Frame Lot No"] = repairedProcess 
            excelData["Process 1 Bushing"] = repairedProcess
            excelData["Process 1 Bushing Lot No"] = repairedProcess
            excelData["Process 1 ST"] = repairedProcess
            excelData["Process 1 Actual Time"] = repairedProcess
            excelData["Process 1 NG Cause"] = repairedProcess
            excelData["Process 1 Repaired Action"] = repairedProcess  

        if process3Status == "Repaired":
            repairedProcess = "REPAIRED AT PROCESS3"
            process3Row += 1
            process4Row += 1
            process5Row += 1
            process6Row += 1
            piRow += 1

            # excelData["DATE"] = repairedProcess
            # excelData["TIME"] = repairedProcess
            # excelData["MODEL CODE"] = repairedProcess
            # excelData["PROCESS S/N"] = repairedProcess
            # excelData["S/N"] = repairedProcess
            # excelData["PASS/NG"] = repairedProcess
            # excelData["VOLTAGE MAX (V)"] = repairedProcess
            # excelData["WATTAGE MAX (W)"] = repairedProcess
            # excelData["CLOSED PRESSURE_MAX (kPa)"] = repairedProcess
            # excelData["VOLTAGE Middle (V)"] = repairedProcess
            # excelData["WATTAGE Middle (W)"] = repairedProcess
            # excelData["AMPERAGE Middle (A)"] = repairedProcess
            # excelData["CLOSED PRESSURE Middle (kPa)"] = repairedProcess
            # excelData["dB(A) 1"] = repairedProcess
            # excelData["dB(A) 2"] = repairedProcess
            # excelData["dB(A) 3"] = repairedProcess
            # excelData["VOLTAGE MIN (V)"] = repairedProcess
            # excelData["WATTAGE MIN (W)"] = repairedProcess
            # excelData["CLOSED PRESSURE MIN (kPa)"] = repairedProcess

            excelData["Process 1 Model Code"] = repairedProcess
            excelData["Process 1 S/N"] = repairedProcess
            excelData["Process 1 ID"] = repairedProcess
            excelData["Process 1 NAME"] = repairedProcess
            excelData["Process 1 Regular/Contractual"] = repairedProcess
            excelData["Process 1 Em2p"] = repairedProcess
            excelData["Process 1 Em2p Lot No"] = repairedProcess
            excelData["Process 1 Em3p"] = repairedProcess
            excelData["Process 1 Em3p Lot No"] = repairedProcess
            excelData["Process 1 Harness"] = repairedProcess
            excelData["Process 1 Harness Lot No"] = repairedProcess
            excelData["Process 1 Frame"] = repairedProcess
            excelData["Process 1 Frame Lot No"] = repairedProcess 
            excelData["Process 1 Bushing"] = repairedProcess
            excelData["Process 1 Bushing Lot No"] = repairedProcess
            excelData["Process 1 ST"] = repairedProcess
            excelData["Process 1 Actual Time"] = repairedProcess
            excelData["Process 1 NG Cause"] = repairedProcess
            excelData["Process 1 Repaired Action"] = repairedProcess  

            excelData["Process 2 Model Code"] = repairedProcess
            excelData["Process 2 S/N"] = repairedProcess
            excelData["Process 2 ID"] = repairedProcess
            excelData["Process 2 NAME"] = repairedProcess
            excelData["Process 2 Regular/Contractual"] = repairedProcess
            excelData["Process 2 M4x40 Screw"] = repairedProcess
            excelData["Process 2 M4x40 Screw Lot No"] = repairedProcess
            excelData["Process 2 Rod Blk"] = repairedProcess
            excelData["Process 2 Rod Blk Lot No"] = repairedProcess
            excelData["Process 2 Df Blk"] = repairedProcess
            excelData["Process 2 Df Blk Lot No"] = repairedProcess
            excelData["Process 2 Df Ring"] = repairedProcess
            excelData["Process 2 Df Ring Lot No"] = repairedProcess
            excelData["Process 2 Washer"] = repairedProcess
            excelData["Process 2 Washer Lot No"] = repairedProcess
            excelData["Process 2 Lock Nut"] = repairedProcess
            excelData["Process 2 Lock Nut Lot No"] = repairedProcess
            excelData["Process 2 ST"] = repairedProcess
            excelData["Process 2 Actual Time"] = repairedProcess
            excelData["Process 2 NG Cause"] = repairedProcess
            excelData["Process 2 Repaired Action"] = repairedProcess

        if process4Status == "Repaired":
            repairedProcess = "REPAIRED AT PROCESS4"
            process4Row += 1
            process5Row += 1
            process6Row += 1
            piRow += 1

            # excelData["DATE"] = repairedProcess
            # excelData["TIME"] = repairedProcess
            # excelData["MODEL CODE"] = repairedProcess
            # excelData["PROCESS S/N"] = repairedProcess
            # excelData["S/N"] = repairedProcess
            # excelData["PASS/NG"] = repairedProcess
            # excelData["VOLTAGE MAX (V)"] = repairedProcess
            # excelData["WATTAGE MAX (W)"] = repairedProcess
            # excelData["CLOSED PRESSURE_MAX (kPa)"] = repairedProcess
            # excelData["VOLTAGE Middle (V)"] = repairedProcess
            # excelData["WATTAGE Middle (W)"] = repairedProcess
            # excelData["AMPERAGE Middle (A)"] = repairedProcess
            # excelData["CLOSED PRESSURE Middle (kPa)"] = repairedProcess
            # excelData["dB(A) 1"] = repairedProcess
            # excelData["dB(A) 2"] = repairedProcess
            # excelData["dB(A) 3"] = repairedProcess
            # excelData["VOLTAGE MIN (V)"] = repairedProcess
            # excelData["WATTAGE MIN (W)"] = repairedProcess
            # excelData["CLOSED PRESSURE MIN (kPa)"] = repairedProcess

            excelData["Process 1 Model Code"] = repairedProcess
            excelData["Process 1 S/N"] = repairedProcess
            excelData["Process 1 ID"] = repairedProcess
            excelData["Process 1 NAME"] = repairedProcess
            excelData["Process 1 Regular/Contractual"] = repairedProcess
            excelData["Process 1 Em2p"] = repairedProcess
            excelData["Process 1 Em2p Lot No"] = repairedProcess
            excelData["Process 1 Em3p"] = repairedProcess
            excelData["Process 1 Em3p Lot No"] = repairedProcess
            excelData["Process 1 Harness"] = repairedProcess
            excelData["Process 1 Harness Lot No"] = repairedProcess
            excelData["Process 1 Frame"] = repairedProcess
            excelData["Process 1 Frame Lot No"] = repairedProcess 
            excelData["Process 1 Bushing"] = repairedProcess
            excelData["Process 1 Bushing Lot No"] = repairedProcess
            excelData["Process 1 ST"] = repairedProcess
            excelData["Process 1 Actual Time"] = repairedProcess
            excelData["Process 1 NG Cause"] = repairedProcess
            excelData["Process 1 Repaired Action"] = repairedProcess  

            excelData["Process 2 Model Code"] = repairedProcess
            excelData["Process 2 S/N"] = repairedProcess
            excelData["Process 2 ID"] = repairedProcess
            excelData["Process 2 NAME"] = repairedProcess
            excelData["Process 2 Regular/Contractual"] = repairedProcess
            excelData["Process 2 M4x40 Screw"] = repairedProcess
            excelData["Process 2 M4x40 Screw Lot No"] = repairedProcess
            excelData["Process 2 Rod Blk"] = repairedProcess
            excelData["Process 2 Rod Blk Lot No"] = repairedProcess
            excelData["Process 2 Df Blk"] = repairedProcess
            excelData["Process 2 Df Blk Lot No"] = repairedProcess
            excelData["Process 2 Df Ring"] = repairedProcess
            excelData["Process 2 Df Ring Lot No"] = repairedProcess
            excelData["Process 2 Washer"] = repairedProcess
            excelData["Process 2 Washer Lot No"] = repairedProcess
            excelData["Process 2 Lock Nut"] = repairedProcess
            excelData["Process 2 Lock Nut Lot No"] = repairedProcess
            excelData["Process 2 ST"] = repairedProcess
            excelData["Process 2 Actual Time"] = repairedProcess
            excelData["Process 2 NG Cause"] = repairedProcess
            excelData["Process 2 Repaired Action"] = repairedProcess

            excelData["Process 3 Model Code"] = repairedProcess
            excelData["Process 3 S/N"] = repairedProcess
            excelData["Process 3 ID"] = repairedProcess
            excelData["Process 3 NAME"] = repairedProcess
            excelData["Process 3 Regular/Contractual"] = repairedProcess
            excelData["Process 3 Frame Gasket"] = repairedProcess
            excelData["Process 3 Frame Gasket Lot No"] = repairedProcess
            excelData["Process 3 Casing Block"] = repairedProcess
            excelData["Process 3 Casing Block Lot No"] = repairedProcess
            excelData["Process 3 Casing Block Inspection 1 Average Data"] = repairedProcess
            excelData["Process 3 Casing Block Inspection 1 Minimum Data"] = repairedProcess
            excelData["Process 3 Casing Block Inspection 1 Maximum Data"] = repairedProcess
            excelData["Process 3 Casing Gasket"] = repairedProcess
            excelData["Process 3 Casing Gasket Lot No"] = repairedProcess
            excelData["Process 3 M4x16 Screw 1"] = repairedProcess
            excelData["Process 3 M4x16 Screw 1 Lot No"] = repairedProcess
            excelData["Process 3 M4x16 Screw 2"] = repairedProcess
            excelData["Process 3 M4x16 Screw 2 Lot No"] = repairedProcess
            excelData["Process 3 Ball Cushion"] = repairedProcess
            excelData["Process 3 Ball Cushion Lot No"] = repairedProcess
            excelData["Process 3 Frame Cover"] = repairedProcess
            excelData["Process 3 Frame Cover Lot No"] = repairedProcess
            excelData["Process 3 Partition Board"] = repairedProcess
            excelData["Process 3 Partition Board Lot No"] = repairedProcess
            excelData["Process 3 Built In Tube 1"] = repairedProcess
            excelData["Process 3 Built In Tube 1 Lot No"] = repairedProcess
            excelData["Process 3 Built In Tube 2"] = repairedProcess
            excelData["Process 3 Built In Tube 2 Lot No"] = repairedProcess
            excelData["Process 3 Head Cover"] = repairedProcess
            excelData["Process 3 Head Cover Lot No"] = repairedProcess
            excelData["Process 3 Casing Packing"] = repairedProcess
            excelData["Process 3 Casing Packing Lot No"] = repairedProcess
            excelData["Process 3 M4x12 Screw"] = repairedProcess
            excelData["Process 3 M4x12 Screw Lot No"] = repairedProcess
            excelData["Process 3 Csb L"] = repairedProcess
            excelData["Process 3 Csb L Lot No"] = repairedProcess
            excelData["Process 3 Csb R"] = repairedProcess
            excelData["Process 3 Csb R Lot No"] = repairedProcess
            excelData["Process 3 Head Packing"] = repairedProcess
            excelData["Process 3 Head Packing Lot No"] = repairedProcess
            excelData["Process 3 ST"] = repairedProcess
            excelData["Process 3 Actual Time"] = repairedProcess
            excelData["Process 3 NG Cause"] = repairedProcess
            excelData["Process 3 Repaired Action"] = repairedProcess

        if process5Status == "Repaired":
            repairedProcess = "RE PI"
            process5Row += 1
            process6Row += 1
            piRow += 1

            # excelData["DATE"] = repairedProcess
            # excelData["TIME"] = repairedProcess
            # excelData["MODEL CODE"] = repairedProcess
            # excelData["PROCESS S/N"] = repairedProcess
            # excelData["S/N"] = repairedProcess
            # excelData["PASS/NG"] = repairedProcess
            # excelData["VOLTAGE MAX (V)"] = repairedProcess
            # excelData["WATTAGE MAX (W)"] = repairedProcess
            # excelData["CLOSED PRESSURE_MAX (kPa)"] = repairedProcess
            # excelData["VOLTAGE Middle (V)"] = repairedProcess
            # excelData["WATTAGE Middle (W)"] = repairedProcess
            # excelData["AMPERAGE Middle (A)"] = repairedProcess
            # excelData["CLOSED PRESSURE Middle (kPa)"] = repairedProcess
            # excelData["dB(A) 1"] = repairedProcess
            # excelData["dB(A) 2"] = repairedProcess
            # excelData["dB(A) 3"] = repairedProcess
            # excelData["VOLTAGE MIN (V)"] = repairedProcess
            # excelData["WATTAGE MIN (W)"] = repairedProcess
            # excelData["CLOSED PRESSURE MIN (kPa)"] = repairedProcess

            excelData["Process 1 Model Code"] = repairedProcess
            excelData["Process 1 S/N"] = repairedProcess
            excelData["Process 1 ID"] = repairedProcess
            excelData["Process 1 NAME"] = repairedProcess
            excelData["Process 1 Regular/Contractual"] = repairedProcess
            excelData["Process 1 Em2p"] = repairedProcess
            excelData["Process 1 Em2p Lot No"] = repairedProcess
            excelData["Process 1 Em3p"] = repairedProcess
            excelData["Process 1 Em3p Lot No"] = repairedProcess
            excelData["Process 1 Harness"] = repairedProcess
            excelData["Process 1 Harness Lot No"] = repairedProcess
            excelData["Process 1 Frame"] = repairedProcess
            excelData["Process 1 Frame Lot No"] = repairedProcess 
            excelData["Process 1 Bushing"] = repairedProcess
            excelData["Process 1 Bushing Lot No"] = repairedProcess
            excelData["Process 1 ST"] = repairedProcess
            excelData["Process 1 Actual Time"] = repairedProcess
            excelData["Process 1 NG Cause"] = repairedProcess
            excelData["Process 1 Repaired Action"] = repairedProcess  

            excelData["Process 2 Model Code"] = repairedProcess
            excelData["Process 2 S/N"] = repairedProcess
            excelData["Process 2 ID"] = repairedProcess
            excelData["Process 2 NAME"] = repairedProcess
            excelData["Process 2 Regular/Contractual"] = repairedProcess
            excelData["Process 2 M4x40 Screw"] = repairedProcess
            excelData["Process 2 M4x40 Screw Lot No"] = repairedProcess
            excelData["Process 2 Rod Blk"] = repairedProcess
            excelData["Process 2 Rod Blk Lot No"] = repairedProcess
            excelData["Process 2 Df Blk"] = repairedProcess
            excelData["Process 2 Df Blk Lot No"] = repairedProcess
            excelData["Process 2 Df Ring"] = repairedProcess
            excelData["Process 2 Df Ring Lot No"] = repairedProcess
            excelData["Process 2 Washer"] = repairedProcess
            excelData["Process 2 Washer Lot No"] = repairedProcess
            excelData["Process 2 Lock Nut"] = repairedProcess
            excelData["Process 2 Lock Nut Lot No"] = repairedProcess
            excelData["Process 2 ST"] = repairedProcess
            excelData["Process 2 Actual Time"] = repairedProcess
            excelData["Process 2 NG Cause"] = repairedProcess
            excelData["Process 2 Repaired Action"] = repairedProcess

            excelData["Process 3 Model Code"] = repairedProcess
            excelData["Process 3 S/N"] = repairedProcess
            excelData["Process 3 ID"] = repairedProcess
            excelData["Process 3 NAME"] = repairedProcess
            excelData["Process 3 Regular/Contractual"] = repairedProcess
            excelData["Process 3 Frame Gasket"] = repairedProcess
            excelData["Process 3 Frame Gasket Lot No"] = repairedProcess
            excelData["Process 3 Casing Block"] = repairedProcess
            excelData["Process 3 Casing Block Lot No"] = repairedProcess
            excelData["Process 3 Casing Block Inspection 1 Average Data"] = repairedProcess
            excelData["Process 3 Casing Block Inspection 1 Minimum Data"] = repairedProcess
            excelData["Process 3 Casing Block Inspection 1 Maximum Data"] = repairedProcess
            excelData["Process 3 Casing Gasket"] = repairedProcess
            excelData["Process 3 Casing Gasket Lot No"] = repairedProcess
            excelData["Process 3 M4x16 Screw 1"] = repairedProcess
            excelData["Process 3 M4x16 Screw 1 Lot No"] = repairedProcess
            excelData["Process 3 M4x16 Screw 2"] = repairedProcess
            excelData["Process 3 M4x16 Screw 2 Lot No"] = repairedProcess
            excelData["Process 3 Ball Cushion"] = repairedProcess
            excelData["Process 3 Ball Cushion Lot No"] = repairedProcess
            excelData["Process 3 Frame Cover"] = repairedProcess
            excelData["Process 3 Frame Cover Lot No"] = repairedProcess
            excelData["Process 3 Partition Board"] = repairedProcess
            excelData["Process 3 Partition Board Lot No"] = repairedProcess
            excelData["Process 3 Built In Tube 1"] = repairedProcess
            excelData["Process 3 Built In Tube 1 Lot No"] = repairedProcess
            excelData["Process 3 Built In Tube 2"] = repairedProcess
            excelData["Process 3 Built In Tube 2 Lot No"] = repairedProcess
            excelData["Process 3 Head Cover"] = repairedProcess
            excelData["Process 3 Head Cover Lot No"] = repairedProcess
            excelData["Process 3 Casing Packing"] = repairedProcess
            excelData["Process 3 Casing Packing Lot No"] = repairedProcess
            excelData["Process 3 M4x12 Screw"] = repairedProcess
            excelData["Process 3 M4x12 Screw Lot No"] = repairedProcess
            excelData["Process 3 Csb L"] = repairedProcess
            excelData["Process 3 Csb L Lot No"] = repairedProcess
            excelData["Process 3 Csb R"] = repairedProcess
            excelData["Process 3 Csb R Lot No"] = repairedProcess
            excelData["Process 3 Head Packing"] = repairedProcess
            excelData["Process 3 Head Packing Lot No"] = repairedProcess
            excelData["Process 3 ST"] = repairedProcess
            excelData["Process 3 Actual Time"] = repairedProcess
            excelData["Process 3 NG Cause"] = repairedProcess
            excelData["Process 3 Repaired Action"] = repairedProcess
            
            excelData["Process 4 Model Code"] = repairedProcess
            excelData["Process 4 S/N"] = repairedProcess
            excelData["Process 4 ID"] = repairedProcess
            excelData["Process 4 NAME"] = repairedProcess
            excelData["Process 4 Regular/Contractual"] = repairedProcess
            excelData["Process 4 Tank"] = repairedProcess
            excelData["Process 4 Tank Lot No"] = repairedProcess
            excelData["Process 4 Upper Housing"] = repairedProcess
            excelData["Process 4 Upper Housing Lot No"] = repairedProcess
            excelData["Process 4 Cord Hook" ] = repairedProcess
            excelData["Process 4 Cord Hook Lot No"] = repairedProcess
            excelData["Process 4 M4x16 Screw"] = repairedProcess
            excelData["Process 4 M4x16 Screw Lot No"] = repairedProcess
            excelData["Process 4 Tank Gasket"] = repairedProcess
            excelData["Process 4 Tank Gasket Lot No"] = repairedProcess
            excelData["Process 4 Tank Cover"] = repairedProcess
            excelData["Process 4 Tank Cover Lot No"] = repairedProcess
            excelData["Process 4 Housing Gasket"] = repairedProcess
            excelData["Process 4 Housing Gasket Lot No"] = repairedProcess
            excelData["Process 4 M4x40 Screw"] = repairedProcess
            excelData["Process 4 M4x40 Screw Lot No"] = repairedProcess
            excelData["Process 4 PartitionGasket"] = repairedProcess
            excelData["Process 4 PartitionGasket Lot No"] = repairedProcess
            excelData["Process 4 M4x12 Screw"] = repairedProcess
            excelData["Process 4 M4x12 Screw Lot No"] = repairedProcess
            excelData["Process 4 Muffler"] = repairedProcess
            excelData["Process 4 Muffler Lot No"] = repairedProcess
            excelData["Process 4 Muffler Gasket"] = repairedProcess
            excelData["Process 4 Muffler Gasket Lot No"] = repairedProcess
            excelData["Process 4 VCR"] = repairedProcess
            excelData["Process 4 VCR Lot No"] = repairedProcess
            excelData["Process 4 ST"] = repairedProcess
            excelData["Process 4 Actual Time"] = repairedProcess
            excelData["Process 4 NG Cause"] = repairedProcess
            excelData["Process 4 Repaired Action"] = repairedProcess

        if process6Status == "Repaired":
            repairedProcess = "REPAIRED AT PROCESS6"
            process6Row += 1
            piRow += 1

            # excelData["DATE"] = repairedProcess
            # excelData["TIME"] = repairedProcess
            # excelData["MODEL CODE"] = repairedProcess
            # excelData["PROCESS S/N"] = repairedProcess
            # excelData["S/N"] = repairedProcess
            # excelData["PASS/NG"] = repairedProcess
            # excelData["VOLTAGE MAX (V)"] = repairedProcess
            # excelData["WATTAGE MAX (W)"] = repairedProcess
            # excelData["CLOSED PRESSURE_MAX (kPa)"] = repairedProcess
            # excelData["VOLTAGE Middle (V)"] = repairedProcess
            # excelData["WATTAGE Middle (W)"] = repairedProcess
            # excelData["AMPERAGE Middle (A)"] = repairedProcess
            # excelData["CLOSED PRESSURE Middle (kPa)"] = repairedProcess
            # excelData["dB(A) 1"] = repairedProcess
            # excelData["dB(A) 2"] = repairedProcess
            # excelData["dB(A) 3"] = repairedProcess
            # excelData["VOLTAGE MIN (V)"] = repairedProcess
            # excelData["WATTAGE MIN (W)"] = repairedProcess
            # excelData["CLOSED PRESSURE MIN (kPa)"] = repairedProcess

            excelData["Process 1 Model Code"] = repairedProcess
            excelData["Process 1 S/N"] = repairedProcess
            excelData["Process 1 ID"] = repairedProcess
            excelData["Process 1 NAME"] = repairedProcess
            excelData["Process 1 Regular/Contractual"] = repairedProcess
            excelData["Process 1 Em2p"] = repairedProcess
            excelData["Process 1 Em2p Lot No"] = repairedProcess
            excelData["Process 1 Em3p"] = repairedProcess
            excelData["Process 1 Em3p Lot No"] = repairedProcess
            excelData["Process 1 Harness"] = repairedProcess
            excelData["Process 1 Harness Lot No"] = repairedProcess
            excelData["Process 1 Frame"] = repairedProcess
            excelData["Process 1 Frame Lot No"] = repairedProcess 
            excelData["Process 1 Bushing"] = repairedProcess
            excelData["Process 1 Bushing Lot No"] = repairedProcess
            excelData["Process 1 ST"] = repairedProcess
            excelData["Process 1 Actual Time"] = repairedProcess
            excelData["Process 1 NG Cause"] = repairedProcess
            excelData["Process 1 Repaired Action"] = repairedProcess  

            excelData["Process 2 Model Code"] = repairedProcess
            excelData["Process 2 S/N"] = repairedProcess
            excelData["Process 2 ID"] = repairedProcess
            excelData["Process 2 NAME"] = repairedProcess
            excelData["Process 2 Regular/Contractual"] = repairedProcess
            excelData["Process 2 M4x40 Screw"] = repairedProcess
            excelData["Process 2 M4x40 Screw Lot No"] = repairedProcess
            excelData["Process 2 Rod Blk"] = repairedProcess
            excelData["Process 2 Rod Blk Lot No"] = repairedProcess
            excelData["Process 2 Df Blk"] = repairedProcess
            excelData["Process 2 Df Blk Lot No"] = repairedProcess
            excelData["Process 2 Df Ring"] = repairedProcess
            excelData["Process 2 Df Ring Lot No"] = repairedProcess
            excelData["Process 2 Washer"] = repairedProcess
            excelData["Process 2 Washer Lot No"] = repairedProcess
            excelData["Process 2 Lock Nut"] = repairedProcess
            excelData["Process 2 Lock Nut Lot No"] = repairedProcess
            excelData["Process 2 ST"] = repairedProcess
            excelData["Process 2 Actual Time"] = repairedProcess
            excelData["Process 2 NG Cause"] = repairedProcess
            excelData["Process 2 Repaired Action"] = repairedProcess

            excelData["Process 3 Model Code"] = repairedProcess
            excelData["Process 3 S/N"] = repairedProcess
            excelData["Process 3 ID"] = repairedProcess
            excelData["Process 3 NAME"] = repairedProcess
            excelData["Process 3 Regular/Contractual"] = repairedProcess
            excelData["Process 3 Frame Gasket"] = repairedProcess
            excelData["Process 3 Frame Gasket Lot No"] = repairedProcess
            excelData["Process 3 Casing Block"] = repairedProcess
            excelData["Process 3 Casing Block Lot No"] = repairedProcess
            excelData["Process 3 Casing Block Inspection 1 Average Data"] = repairedProcess
            excelData["Process 3 Casing Block Inspection 1 Minimum Data"] = repairedProcess
            excelData["Process 3 Casing Block Inspection 1 Maximum Data"] = repairedProcess
            excelData["Process 3 Casing Gasket"] = repairedProcess
            excelData["Process 3 Casing Gasket Lot No"] = repairedProcess
            excelData["Process 3 M4x16 Screw 1"] = repairedProcess
            excelData["Process 3 M4x16 Screw 1 Lot No"] = repairedProcess
            excelData["Process 3 M4x16 Screw 2"] = repairedProcess
            excelData["Process 3 M4x16 Screw 2 Lot No"] = repairedProcess
            excelData["Process 3 Ball Cushion"] = repairedProcess
            excelData["Process 3 Ball Cushion Lot No"] = repairedProcess
            excelData["Process 3 Frame Cover"] = repairedProcess
            excelData["Process 3 Frame Cover Lot No"] = repairedProcess
            excelData["Process 3 Partition Board"] = repairedProcess
            excelData["Process 3 Partition Board Lot No"] = repairedProcess
            excelData["Process 3 Built In Tube 1"] = repairedProcess
            excelData["Process 3 Built In Tube 1 Lot No"] = repairedProcess
            excelData["Process 3 Built In Tube 2"] = repairedProcess
            excelData["Process 3 Built In Tube 2 Lot No"] = repairedProcess
            excelData["Process 3 Head Cover"] = repairedProcess
            excelData["Process 3 Head Cover Lot No"] = repairedProcess
            excelData["Process 3 Casing Packing"] = repairedProcess
            excelData["Process 3 Casing Packing Lot No"] = repairedProcess
            excelData["Process 3 M4x12 Screw"] = repairedProcess
            excelData["Process 3 M4x12 Screw Lot No"] = repairedProcess
            excelData["Process 3 Csb L"] = repairedProcess
            excelData["Process 3 Csb L Lot No"] = repairedProcess
            excelData["Process 3 Csb R"] = repairedProcess
            excelData["Process 3 Csb R Lot No"] = repairedProcess
            excelData["Process 3 Head Packing"] = repairedProcess
            excelData["Process 3 Head Packing Lot No"] = repairedProcess
            excelData["Process 3 ST"] = repairedProcess
            excelData["Process 3 Actual Time"] = repairedProcess
            excelData["Process 3 NG Cause"] = repairedProcess
            excelData["Process 3 Repaired Action"] = repairedProcess
            
            excelData["Process 4 Model Code"] = repairedProcess
            excelData["Process 4 S/N"] = repairedProcess
            excelData["Process 4 ID"] = repairedProcess
            excelData["Process 4 NAME"] = repairedProcess
            excelData["Process 4 Regular/Contractual"] = repairedProcess
            excelData["Process 4 Tank"] = repairedProcess
            excelData["Process 4 Tank Lot No"] = repairedProcess
            excelData["Process 4 Upper Housing"] = repairedProcess
            excelData["Process 4 Upper Housing Lot No"] = repairedProcess
            excelData["Process 4 Cord Hook" ] = repairedProcess
            excelData["Process 4 Cord Hook Lot No"] = repairedProcess
            excelData["Process 4 M4x16 Screw"] = repairedProcess
            excelData["Process 4 M4x16 Screw Lot No"] = repairedProcess
            excelData["Process 4 Tank Gasket"] = repairedProcess
            excelData["Process 4 Tank Gasket Lot No"] = repairedProcess
            excelData["Process 4 Tank Cover"] = repairedProcess
            excelData["Process 4 Tank Cover Lot No"] = repairedProcess
            excelData["Process 4 Housing Gasket"] = repairedProcess
            excelData["Process 4 Housing Gasket Lot No"] = repairedProcess
            excelData["Process 4 M4x40 Screw"] = repairedProcess
            excelData["Process 4 M4x40 Screw Lot No"] = repairedProcess
            excelData["Process 4 PartitionGasket"] = repairedProcess
            excelData["Process 4 PartitionGasket Lot No"] = repairedProcess
            excelData["Process 4 M4x12 Screw"] = repairedProcess
            excelData["Process 4 M4x12 Screw Lot No"] = repairedProcess
            excelData["Process 4 Muffler"] = repairedProcess
            excelData["Process 4 Muffler Lot No"] = repairedProcess
            excelData["Process 4 Muffler Gasket"] = repairedProcess
            excelData["Process 4 Muffler Gasket Lot No"] = repairedProcess
            excelData["Process 4 VCR"] = repairedProcess
            excelData["Process 4 VCR Lot No"] = repairedProcess
            excelData["Process 4 ST"] = repairedProcess
            excelData["Process 4 Actual Time"] = repairedProcess
            excelData["Process 4 NG Cause"] = repairedProcess
            excelData["Process 4 Repaired Action"] = repairedProcess

            excelData["Process 5 Model Code"] = repairedProcess
            excelData["Process 5 S/N"] = repairedProcess
            excelData["Process 5 ID"] = repairedProcess
            excelData["Process 5 NAME"] = repairedProcess
            excelData["Process 5 Regular/Contractual"] = repairedProcess
            excelData["Process 5 Rating Label"] = repairedProcess
            excelData["Process 5 Rating Label Lot No"] = repairedProcess
            excelData["Process 5 ST"] = repairedProcess
            excelData["Process 5 Actual Time"] = repairedProcess
            excelData["Process 5 NG Cause"] = repairedProcess
            excelData["Process 5 Repaired Action"] = repairedProcess
        
    compiledFrame = pd.concat([compiledFrame, excelData], ignore_index=True)

# %%
def StopProgram():
    global programRunning
    global autoRun
    programRunning = False
    autoRun = False

    root.destroy()

# %%
def start():
    global compiledFrame
    global canCompile
    global isCsvReaded
    global readCount
    global programRunning
    global process1Row
    global process2Row
    global process3Row
    global process4Row
    global process5Row
    global process6Row
    global piRow

    #GUI
    global compileButton
    global loadingText

    process1Row = 0
    process2Row = 0
    process3Row = 0
    process4Row = 0
    process5Row = 0
    process6Row = 0
    piRow = 0
    programRunning = True
    isCsvReaded = False

    setDate()

    compileButton.config(text= loadingText)
    compileButton.config(state= "disabled")

    col = [
            "DATE", 
            "TIME", 
            "MODEL CODE", 
            "PROCESS S/N", 
            "S/N", 
            "PASS/NG", 
            "VOLTAGE MAX (V)", 
            "WATTAGE MAX (W)", 
            "CLOSED PRESSURE_MAX (kPa)", 
            "VOLTAGE Middle (V)", 
            "WATTAGE Middle (W)", 
            "AMPERAGE Middle (A)", 
            "CLOSED PRESSURE Middle (kPa)", 
            "dB(A) 1", 
            "dB(A) 2", 
            "dB(A) 3", 
            "VOLTAGE MIN (V)", 
            "WATTAGE MIN (W)", 
            "CLOSED PRESSURE MIN (kPa)",

            "Process 1 Model Code",
            "Process 1 S/N",
            "Process 1 ID",
            "Process 1 NAME",
            "Process 1 Regular/Contractual",
            "Process 1 Em2p",
            "Process 1 Em2p Lot No",
            "Process 1 Em2p Inspection 3 Average Data",
            "Process 1 Em2p Inspection 4 Average Data",
            "Process 1 Em2p Inspection 5 Average Data",
            "Process 1 Em2p Inspection 10 Average Data",
            "Process 1 Em2p Inspection 3 Minimum Data",
            "Process 1 Em2p Inspection 4 Minimum Data",
            "Process 1 Em2p Inspection 5 Minimum Data",
            "Process 1 Em2p Inspection 3 Maximum Data",
            "Process 1 Em2p Inspection 4 Maximum Data",
            "Process 1 Em2p Inspection 5 Maximum Data",
            "Process 1 Em3p",
            "Process 1 Em3p Lot No",
            "Process 1 Em3p Inspection 3 Average Data",
            "Process 1 Em3p Inspection 4 Average Data",
            "Process 1 Em3p Inspection 5 Average Data",
            "Process 1 Em3p Inspection 10 Average Data",
            "Process 1 Em3p Inspection 3 Minimum Data",
            "Process 1 Em3p Inspection 4 Minimum Data",
            "Process 1 Em3p Inspection 5 Minimum Data",
            "Process 1 Em3p Inspection 3 Maximum Data",
            "Process 1 Em3p Inspection 4 Maximum Data",
            "Process 1 Em3p Inspection 5 Maximum Data",
            "Process 1 Harness",
            "Process 1 Harness Lot No",
            "Process 1 Frame",
            "Process 1 Frame Lot No",
            "Process 1 Frame Inspection 1 Average Data", 
            "Process 1 Frame Inspection 2 Average Data", 
            "Process 1 Frame Inspection 3 Average Data", 
            "Process 1 Frame Inspection 4 Average Data", 
            "Process 1 Frame Inspection 5 Average Data", 
            "Process 1 Frame Inspection 6 Average Data", 
            "Process 1 Frame Inspection 7 Average Data", 
            "Process 1 Frame Inspection 1 Minimum Data", 
            "Process 1 Frame Inspection 2 Minimum Data", 
            "Process 1 Frame Inspection 3 Minimum Data", 
            "Process 1 Frame Inspection 4 Minimum Data", 
            "Process 1 Frame Inspection 5 Minimum Data", 
            "Process 1 Frame Inspection 6 Minimum Data", 
            "Process 1 Frame Inspection 7 Minimum Data", 
            "Process 1 Frame Inspection 1 Maximum Data", 
            "Process 1 Frame Inspection 2 Maximum Data", 
            "Process 1 Frame Inspection 3 Maximum Data", 
            "Process 1 Frame Inspection 4 Maximum Data", 
            "Process 1 Frame Inspection 5 Maximum Data", 
            "Process 1 Frame Inspection 6 Maximum Data", 
            "Process 1 Frame Inspection 7 Maximum Data", 
            "Process 1 Bushing",
            "Process 1 Bushing Lot No",
            "Process 1 ST",
            "Process 1 Actual Time",
            "Process 1 NG Cause",
            "Process 1 Repaired Action",

            "Process 2 Model Code",
            "Process 2 S/N",
            "Process 2 ID",
            "Process 2 NAME",
            "Process 2 Regular/Contractual",
            "Process 2 M4x40 Screw",
            "Process 2 M4x40 Screw Lot No",
            "Process 2 Rod Blk",
            "Process 2 Rod Blk Lot No",
            "Process 2 Rod Blk Tesla 1 Average Data",
            "Process 2 Rod Blk Tesla 2 Average Data",
            "Process 2 Rod Blk Tesla 3 Average Data",
            "Process 2 Rod Blk Tesla 4 Average Data",
            "Process 2 Rod Blk Tesla 1 Minimum Data",
            "Process 2 Rod Blk Tesla 2 Minimum Data",
            "Process 2 Rod Blk Tesla 3 Minimum Data",
            "Process 2 Rod Blk Tesla 4 Minimum Data",
            "Process 2 Rod Blk Tesla 1 Maximum Data",
            "Process 2 Rod Blk Tesla 2 Maximum Data",
            "Process 2 Rod Blk Tesla 3 Maximum Data",
            "Process 2 Rod Blk Tesla 4 Maximum Data",
            "Process 2 Rod Blk Inspection 1 Average Data",
            "Process 2 Rod Blk Inspection 2 Average Data",
            "Process 2 Rod Blk Inspection 3 Average Data",
            "Process 2 Rod Blk Inspection 4 Average Data",
            "Process 2 Rod Blk Inspection 5 Average Data",
            "Process 2 Rod Blk Inspection 6 Average Data",
            "Process 2 Rod Blk Inspection 7 Average Data",
            "Process 2 Rod Blk Inspection 8 Average Data",
            "Process 2 Rod Blk Inspection 9 Average Data",
            "Process 2 Rod Blk Inspection 1 Minimum Data",
            "Process 2 Rod Blk Inspection 2 Minimum Data",
            "Process 2 Rod Blk Inspection 3 Minimum Data",
            "Process 2 Rod Blk Inspection 4 Minimum Data",
            "Process 2 Rod Blk Inspection 5 Minimum Data",
            "Process 2 Rod Blk Inspection 6 Minimum Data",
            "Process 2 Rod Blk Inspection 7 Minimum Data",
            "Process 2 Rod Blk Inspection 8 Minimum Data",
            "Process 2 Rod Blk Inspection 9 Minimum Data",
            "Process 2 Rod Blk Inspection 1 Maximum Data",
            "Process 2 Rod Blk Inspection 2 Maximum Data",
            "Process 2 Rod Blk Inspection 3 Maximum Data",
            "Process 2 Rod Blk Inspection 4 Maximum Data",
            "Process 2 Rod Blk Inspection 5 Maximum Data",
            "Process 2 Rod Blk Inspection 6 Maximum Data",
            "Process 2 Rod Blk Inspection 7 Maximum Data",
            "Process 2 Rod Blk Inspection 8 Maximum Data",
            "Process 2 Rod Blk Inspection 9 Maximum Data",
            "Process 2 Df Blk",
            "Process 2 Df Blk Lot No",
            "Process 2 Df Blk Inspection 1 Average Data",
            "Process 2 Df Blk Inspection 2 Average Data",
            "Process 2 Df Blk Inspection 3 Average Data",
            "Process 2 Df Blk Inspection 4 Average Data",
            "Process 2 Df Blk Inspection 1 Minimum Data",
            "Process 2 Df Blk Inspection 2 Minimum Data",
            "Process 2 Df Blk Inspection 3 Minimum Data",
            "Process 2 Df Blk Inspection 4 Minimum Data",
            "Process 2 Df Blk Inspection 1 Maximum Data",
            "Process 2 Df Blk Inspection 2 Maximum Data",
            "Process 2 Df Blk Inspection 3 Maximum Data",
            "Process 2 Df Blk Inspection 4 Maximum Data",
            "Process 2 Df Blk Tensile Rate Of Change Average",
            "Process 2 Df Blk Tensile Rate Of Change Minimum",
            "Process 2 Df Blk Tensile Rate Of Change Maximum",
            "Process 2 Df Blk Tensile Start Force Average",
            "Process 2 Df Blk Tensile Start Force Minimum",
            "Process 2 Df Blk Tensile Start Force Maximum",
            "Process 2 Df Blk Tensile Terminating Force Average",
            "Process 2 Df Blk Tensile Terminating Force Minimum",
            "Process 2 Df Blk Tensile Terminating Force Maximum",
            "Process 2 Df Ring",
            "Process 2 Df Ring Lot No",
            "Process 2 Washer",
            "Process 2 Washer Lot No",
            "Process 2 Lock Nut",
            "Process 2 Lock Nut Lot No",
            "Process 2 ST",
            "Process 2 Actual Time",
            "Process 2 NG Cause",
            "Process 2 Repaired Action",

            "Process 3 Model Code",
            "Process 3 S/N",
            "Process 3 ID",
            "Process 3 NAME",
            "Process 3 Regular/Contractual",
            "Process 3 Frame Gasket",
            "Process 3 Frame Gasket Lot No",
            "Process 3 Casing Block",
            "Process 3 Casing Block Lot No",
            "Process 3 Casing Block Inspection 1 Average Data",
            "Process 3 Casing Block Inspection 1 Minimum Data",
            "Process 3 Casing Block Inspection 1 Maximum Data",
            "Process 3 Casing Gasket",
            "Process 3 Casing Gasket Lot No",
            "Process 3 M4x16 Screw 1",
            "Process 3 M4x16 Screw 1 Lot No",
            "Process 3 M4x16 Screw 2",
            "Process 3 M4x16 Screw 2 Lot No",
            "Process 3 Ball Cushion",
            "Process 3 Ball Cushion Lot No",
            "Process 3 Frame Cover",
            "Process 3 Frame Cover Lot No",
            "Process 3 Partition Board",
            "Process 3 Partition Board Lot No",
            "Process 3 Built In Tube 1",
            "Process 3 Built In Tube 1 Lot No",
            "Process 3 Built In Tube 2",
            "Process 3 Built In Tube 2 Lot No",
            "Process 3 Head Cover",
            "Process 3 Head Cover Lot No",
            "Process 3 Casing Packing",
            "Process 3 Casing Packing Lot No",
            "Process 3 M4x12 Screw",
            "Process 3 M4x12 Screw Lot No",
            "Process 3 Csb L",
            "Process 3 Csb L Lot No",
            "Process 3 Csb R",
            "Process 3 Csb R Lot No",
            "Process 3 Head Packing",
            "Process 3 Head Packing Lot No",
            "Process 3 ST",
            "Process 3 Actual Time",
            "Process 3 NG Cause",
            "Process 3 Repaired Action",

            "Process 4 Model Code",
            "Process 4 S/N",
            "Process 4 ID",
            "Process 4 NAME",
            "Process 4 Regular/Contractual",
            "Process 4 Tank",
            "Process 4 Tank Lot No",
            "Process 4 Upper Housing",
            "Process 4 Upper Housing Lot No",
            "Process 4 Cord Hook",
            "Process 4 Cord Hook Lot No",
            "Process 4 M4x16 Screw",
            "Process 4 M4x16 Screw Lot No",
            "Process 4 Tank Gasket",
            "Process 4 Tank Gasket Lot No",
            "Process 4 Tank Cover",
            "Process 4 Tank Cover Lot No",
            "Process 4 Housing Gasket",
            "Process 4 Housing Gasket Lot No",
            "Process 4 M4x40 Screw",
            "Process 4 M4x40 Screw Lot No",
            "Process 4 PartitionGasket",
            "Process 4 PartitionGasket Lot No",
            "Process 4 M4x12 Screw",
            "Process 4 M4x12 Screw Lot No",
            "Process 4 Muffler",
            "Process 4 Muffler Lot No",
            "Process 4 Muffler Gasket",
            "Process 4 Muffler Gasket Lot No",
            "Process 4 VCR",
            "Process 4 VCR Lot No",
            "Process 4 ST",
            "Process 4 Actual Time",
            "Process 4 NG Cause",
            "Process 4 Repaired Action",
            
            "Process 5 Model Code",
            "Process 5 S/N",
            "Process 5 ID",
            "Process 5 NAME",
            "Process 5 Regular/Contractual",
            "Process 5 Rating Label",
            "Process 5 Rating Label Lot No",
            "Process 5 ST",
            "Process 5 Actual Time",
            "Process 5 NG Cause",
            "Process 5 Repaired Action",
            
            "Process 6 Model Code",
            "Process 6 S/N",
            "Process 6 ID",
            "Process 6 NAME",
            "Process 6 Regular/Contractual",
            "Process 6 Vinyl",
            "Process 6 Vinyl Lot No",
            "Process 6 ST",
            "Process 6 Actual Time",
            "Process 6 NG Cause",
            "Process 6 Repaired Action"
        ]
    compiledFrame = pd.DataFrame(columns=col)

    DateAndTimeManager.GetDateToday()

    #READING ALL FILES USING FILES READER
    filesreader = filesReader()
    filesreader.readingYearStored = DateAndTimeManager.yearNow
    filesreader.ReadEm2pFiles()
    filesreader.ReadEm3pFiles()
    filesreader.ReadFmFiles()
    filesreader.ReadDfbSnapFiles()
    filesreader.ReadDfbFiles()
    filesreader.ReadTensile()
    filesreader.ReadRdbCheckSheetFiles()
    filesreader.ReadRdbFiles()
    filesreader.ReadCsbFiles()

    try:
        #Checking If There's Master Pump Data
        CheckPICsv()

        #Writing Master Pump Data
        if canCompilePI:
            CompilePICsv()
            WriteCsv(compiledFrame)

        #Reading VT CSV Files
        while not isCsvReaded:
            try:
                ReadCsv()
                isCsvReaded = True
            except:
                print("Cannot Read Csv Retrying In 1 Seconds")
                isCsvReaded = False
                time.sleep(1)


        # !--------- UNUSED CODEBLOCK ---------!
        #Getting VT Original File
        process1OrigFile = os.path.getmtime(r'\\192.168.2.10\csv\csv\VT1\log000_1.csv')
        process2OrigFile = os.path.getmtime(r'\\192.168.2.10\csv\csv\VT2\log000_2.csv')
        process3OrigFile = os.path.getmtime(r'\\192.168.2.10\csv\csv\VT3\log000_3.csv')
        process4OrigFile = os.path.getmtime(r'\\192.168.2.10\csv\csv\VT4\log000_4.csv')
        process5OrigFile = os.path.getmtime(r'\\192.168.2.10\csv\csv\VT5\log000_5.csv')
        process6OrigFile = os.path.getmtime(r'\\192.168.2.10\csv\csv\VT6\log000_6.csv')

        piOrigFile = os.path.getmtime(r'\\192.168.2.19\ai_team\AI Program\Outputs\PICompiled6.csv')

        # !--------- UNUSED CODEBLOCK ---------!

        while programRunning:
            CsvOrganize()
            if canCompile:
                CompileCsv()
            if loadingText == "Loading...":
                loadingText = "Loading"
            else:
                loadingText += "."
                compileButton.config(text= loadingText)

            # print(f"Process 1 S/N{tempDfVt1["Process 1 S/N"]}_______________________________________________________________________________")
            # print(f"Process 2 S/N{tempDfVt2["Process 2 S/N"]}_______________________________________________________________________________")
            # print(f"Process 3 S/N{tempDfVt3["Process 3 S/N"]}_______________________________________________________________________________")
            # print(f"Process 4 S/N{tempDfVt4["Process 4 S/N"]}_______________________________________________________________________________")
            # print(f"Process 5 S/N{tempDfVt5["Process 5 S/N"]}_______________________________________________________________________________")
            # print(f"Process 6 S/N{tempDfVt6["Process 6 S/N"]}_______________________________________________________________________________")

            #Clearing Cmd Logs When Reaches 10 Lines
            readCount += 1
            if readCount >= 10:
                os.system('cls')
                readCount = 0
            #_______________________________________
        WriteCsv(compiledFrame)
        compileButton.config(text= "Done")
        openOutputDirectory()
        time.sleep(2)
        compileButton.config(text= "COMPILE")
        compileButton.config(state= "normal")
    except FileNotFoundError:
        ErrorReading()
        compileButton.config(text= "COMPILE")
        compileButton.config(state= "normal")

# %%
def StartProgram():
    threading.Thread(target=start).start()

# %%
def openOutputDirectory():
    location = r'\\192.168.2.19\ai_team\AI Program\Outputs\CompiledProcess'
    os.startfile(location)

# %%
def AutoRun():
    global autoRun
    global time_picker

    coolDown = False

    while autoRun:
        print("Auto Run Activated")
        DateAndTimeManager.GetTimeNow()
        print(DateAndTimeManager.timeNow)
        
        hour = time_picker.hours()
        minutes = time_picker.minutes()
        period = time_picker.period()

        timeSet = f"{hour}:{minutes} {period}"
        timeSet = datetime2.strptime(timeSet, "%I:%M %p")
        timeSet = timeSet.strftime("%H:%M")

        print(timeSet)

        if DateAndTimeManager.timeNow == timeSet and not coolDown:
            coolDown = True
            StartProgram()
            time.sleep(70)
            coolDown = False
        time.sleep(1)


# %%
def toggleAutoRun():
    global autoRunButton
    global autoRun
    global compileButton
    global calendarPicker

    if not autoRun:
        autoRunButton.config(image = on)
        compileButton.config(state = "disabled")
        calendarPicker.config(state = "disabled")
        autoRun = True
        setDate()
        AutoRun()
    else:
        autoRunButton.config(image = off)
        compileButton.config(state = "normal")
        calendarPicker.config(state = "normal")
        autoRun = False
        setDate()

# %%
def ToggleAutoRun():
    threading.Thread(target=toggleAutoRun).start()

# %%
def Configure():
    global frame1
    global frame2

    frame1.pack_forget()
    frame2.pack()

# %%
def Back():
    global frame1
    global frame2

    frame1.pack()
    frame2.pack_forget()

# %%
#Ok/Cancel Dialog
def SetTimeConfirmation():
    answer = askokcancel(title='Confirmation', message='Are you sure you want to set the time?')
    if answer:
        Back()

# %%
def setDate():
    global dateToRead
    global dateToReadDashFormat

    global calendarPicker
    global autoRun

    if autoRun:
        DateAndTimeManager.GetDateToday()
        dateToRead = DateAndTimeManager.dateToday
        dateToReadDashFormat = dateToRead.replace("/", "-")
        # QualityControlDataManager.yearSet = dateToReadDashFormat[:4]
        print(f"Date To Read: {dateToReadDashFormat}")
        # print(f"Year To Read: {QualityControlDataManager.yearSet}")
    else:
        selectedDate = calendarPicker.get_date()
        selectedDate = selectedDate.strftime("%Y/%m/%d")

        dateToRead = selectedDate
        dateToReadDashFormat = dateToRead.replace("/", "-")
        # QualityControlDataManager.yearSet = dateToReadDashFormat[:4]
        print(f"Date To Read: {dateToReadDashFormat}")
        # print(f"Year To Read: {QualityControlDataManager.yearSet}")
        

# %%
def ErrorReading():
    showerror(title='Deletion Status', message='File Not Found')

# %%
#Fixing Blur
windll.shcore.SetProcessDpiAwareness(1)

root = tk.Tk()
root.title('FC1 Compiler')
root.iconbitmap('Icons/HiblowLogo.ico')
root.geometry('600x650+50+50')
root.resizable(False, False)

on = PhotoImage(file = "Icons/on.png")
off = PhotoImage(file = "Icons/off.png")

#Frames
frame1 = tk.Frame(root)
frame1.pack()

frame2 = tk.Frame(root)
frame2.pack_forget()

# configure the grid
frame1.columnconfigure(0, weight=1)
frame1.columnconfigure(1, weight=1)

#FRAME1

# place a label on the root window
message = tk.Label(frame1, text="FC1 Compiler", font=("Arial", 12, "bold"))
message.grid(column=0, row=0, columnspan=2, padx=220)

# button
compileButton = tk.Button(frame1, text='COMPILE', font=("Arial", 12), command = StartProgram, width=15, height=1)
compileButton.grid(column=0, row=1, ipadx=5, ipady=5, pady=10, columnspan=2)
compileButton.config(bg="lightgreen", fg="black")

autoRunLabel = tk.Label(frame1, text="Auto Run", font=("Arial", 12, "bold"))
autoRunLabel.grid(column=0, row=2)

autoRunButton = tk.Button(frame1, image = off, bd = 0, font=("Arial", 12), command = ToggleAutoRun)
autoRunButton.grid(column=1, row=2, ipadx=5, ipady=5, pady=10)

configureButton = tk.Button(frame1, text='CONFIGURE', font=("Arial", 8), command = Configure, width=10, height=1)
configureButton.grid(column=0, row=3, ipadx=5, ipady=5, pady=10, columnspan=2)
configureButton.config(bg="lightgreen", fg="black")

calendarPicker = DateEntry(frame1, width=16, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy/mm/dd')
calendarPicker.grid(column=0, row=4, pady=20, padx=(200, 0))

#FRAME2

message = tk.Label(frame2, text="Configure", font=("Arial", 12, "bold"))
message.grid(column=0, row=1, columnspan=2, padx=220)

# button
backButton = tk.Button(frame2, text='BACK', font=("Arial", 8), command = Back, width=10, height=1)
backButton.grid(column=0, row=0, ipadx=5, ipady=5, sticky=W)
backButton.config(bg="lightgreen", fg="black")

applyButton = tk.Button(frame2, text='APPLY', font=("Arial", 8), command = SetTimeConfirmation, width=10, height=1)
applyButton.grid(column=1, row=0, ipadx=5, ipady=5, sticky=E)
applyButton.config(bg="lightgreen", fg="black")

time_picker = AnalogPicker(frame2)
time_picker.grid(column = 0, row = 4, columnspan = 2)
theme = AnalogThemes(time_picker)
theme.setNavyBlue()

# StartProgram()

root.protocol("WM_DELETE_WINDOW", StopProgram)
root.mainloop()