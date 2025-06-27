from Imports import *
import DateAndTimeManager

dfPi = ""
dfPiNotDone = []
tempdfPi = ""
canCompilePI = False
compiledFrame = ""

piRow = 0

# %%
def CheckPICsv():
    global dfPi
    global dfPiNotDone
    global canCompilePI

    global piRow

    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)
    
    piDirectory = (r'\\192.168.2.19\ai_team\AI Program\Outputs\PICompiled')
    os.chdir(piDirectory)

    dfPi = pd.read_csv(f'PICompiled{DateAndTimeManager.dateToReadDashFormat}.csv', encoding='latin1')
    
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
                    "Process 6 Repaired Action" : [processData],

                    "Process 1 SERIAL NO" : [processData],
                    "Process 2 SERIAL NO" : [processData],
                    "Process 3 SERIAL NO" : [processData],
                    "Process 4 SERIAL NO" : [processData],
                    "Process 5 SERIAL NO" : [processData],
                    "Process 6 SERIAL NO" : [processData]
                }
        excelData2 = pd.DataFrame(excelData2)
        compiledFrame = pd.concat([compiledFrame, excelData2], ignore_index=True)

    canCompilePI = False

def ResetVariables():
    global dfPi
    global dfPiNotDone
    global tempdfPi
    global canCompilePI
    global compiledFrame

    global piRow

    dfPi = ""
    dfPiNotDone = []
    tempdfPi = ""
    canCompilePI = False
    compiledFrame = ""

    piRow = 0
