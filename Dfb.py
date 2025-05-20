#%%
from Imports import *
import DateAndTimeManager
from FilesReader import *

#%%
class dFB():
    dfbSnapData = ""
    dfbLetterCode = ""
    dfbLotNumber = ""
    dfbMonth = ""
    dfbCode = ""
    dfbLotNumber2 = ""
    dfbYear = ""

    totalAverage1 = []
    totalAverage2 = []
    totalAverage3 = []
    totalAverage4 = []

    totalMinimum1 = []
    totalMinimum2 = []
    totalMinimum3 = []
    totalMinimum4 = []

    totalMaximum1 = []
    totalMaximum2 = []
    totalMaximum3 = []
    totalMaximum4 = []

    readingYear = ""

    fileList = []
    fileFinishedReading = False

    def __init__(self):
        pass
    def ReadDfbSnap(self, lotNumber):
        self.dfbLetterCode = lotNumber[-1]
        self.dfbYear = lotNumber[:-6]

        #Removing The Last Two Values Of Lot Number
        lotNumber = lotNumber[:-2]
        #Changing The Format Of Lot Number
        lotNumber = datetime2.strptime(lotNumber, "%Y%m%d")
        self.dfbMonth = lotNumber.strftime("%B")
        self.dfbLotNumber = lotNumber.strftime("%Y-%m-%d")
        

        pd.set_option('display.max_columns', None)
        pd.set_option('display.max_rows', None)

        for fileNum in range(len(DFBSNAPData)):
            print(fileNum)
            try:
                #Checking SNAPDATA Per Sheets
                for s in DFBSNAPData[fileNum].sheetnames:
                    if self.dfbMonth.lower() in s.lower():
                        sheet = DFBSNAPData[fileNum][s]
                        self.dfbSnapData = pd.DataFrame(sheet.values)
                        self.dfbSnapData = self.dfbSnapData.iloc[6:]
                        self.dfbSnapData = self.dfbSnapData.replace(r'\s+', '', regex=True)

                        #Getting DFB Code
                        self.dfbCode = self.dfbSnapData.iloc[1, 3]
                        self.dfbCode = self.dfbCode[8:]
                        self.dfbCode = self.dfbCode[:-28]

                        #Filtering SNAP Data, That Contains DFB6600600
                        self.dfbSnapData = self.dfbSnapData[(self.dfbSnapData[1].isin(["DFB6600600"]))]

                        break



                #Converting The First Column/Date To String
                self.dfbSnapData.iloc[:, 0] = self.dfbSnapData.iloc[:, 0].astype(str)

                tempDfbSnapData = self.dfbSnapData[(self.dfbSnapData[0].isin([f"{self.dfbLotNumber} 00:00:00"])) & (self.dfbSnapData[2].isin([self.dfbLetterCode]))]

                self.dfbLotNumber2 = tempDfbSnapData.iloc[:,3].values[0]

                print(f"Dfb Code {self.dfbCode}")
                print(f"Dfb Lot Number {self.dfbLotNumber2}")

                break

            except:
                print('No DFB6600600 Snap Not Found')


        print(self.dfbLotNumber)
        print(self.dfbMonth)
        print(self.dfbLetterCode)

    def GettingData(self, itemCode):
        if itemCode == "DFB6600600":
            self.fileList = DF06600600Data

        for fileNum in range(len(self.fileList)):
            self.totalAverage1 = []
            self.totalAverage2 = []
            self.totalAverage3 = []
            self.totalAverage4 = []

            self.totalMinimum1 = []
            self.totalMinimum2 = []
            self.totalMinimum3 = []
            self.totalMinimum4 = []

            self.totalMaximum1 = []
            self.totalMaximum2 = []
            self.totalMaximum3 = []
            self.totalMaximum4 = []

            print(f"READING FILE {fileNum}")

            #Getting The Row, Column Location Of SUPPLIER
            findSupplier = [(index, column) for index, row in self.fileList[fileNum].iterrows() for column, value in row.items() if value == "SUPPLIER"]
            supplierRow = [index for index, _ in findSupplier]
            supplierColumn = [column for _, column in findSupplier]

            print("Row indices:", supplierRow)
            print("Column names:", supplierColumn)

            # Get the Neighboring Data Of Supplier
            supplierFiltered = self.fileList[fileNum].iloc[max(0, supplierRow[0] - 3):min(len(self.fileList[fileNum]), supplierRow[0] + 10), self.fileList[fileNum].columns.get_loc(supplierColumn[0]):self.fileList[fileNum].columns.get_loc(supplierColumn[0]) + 999999]
            
            try:
                #Getting The Row, Column Location Of Lot Number
                findLotNumber = [(index, column) for index, row in supplierFiltered.iterrows() for column, value in row.items() if value == self.dfbLotNumber2[:-3]]
                lotNumberRow = [index for index, _ in findLotNumber]
                lotNumberColumn = [column for _, column in findLotNumber]

                print("Row indices:", lotNumberRow)
                print("Column names:", lotNumberColumn)

                for a in range(0, len(lotNumberColumn)):
                    inspectionData = self.fileList[fileNum].iloc[max(0, lotNumberRow[a]):min(len(self.fileList[fileNum]), lotNumberRow[a] + 10), self.fileList[fileNum].columns.get_loc(lotNumberColumn[a]):self.fileList[fileNum].columns.get_loc(lotNumberColumn[a]) + 5]

                    average1 = inspectionData.iloc[3].mean()
                    average2 = inspectionData.iloc[4].mean()
                    average3 = inspectionData.iloc[5].mean()
                    average4 = inspectionData.iloc[6].mean()

                    minimum1 = inspectionData.iloc[3].min()
                    minimum2 = inspectionData.iloc[4].min()
                    minimum3 = inspectionData.iloc[5].min()
                    minimum4 = inspectionData.iloc[6].min()

                    maximum1 = inspectionData.iloc[3].max()
                    maximum2 = inspectionData.iloc[4].max()
                    maximum3 = inspectionData.iloc[5].max()
                    maximum4 = inspectionData.iloc[6].max()

                    self.totalAverage1.append(average1)
                    self.totalAverage2.append(average2)
                    self.totalAverage3.append(average3)
                    self.totalAverage4.append(average4)

                    self.totalMinimum1.append(minimum1)
                    self.totalMinimum2.append(minimum2)
                    self.totalMinimum3.append(minimum3)
                    self.totalMinimum4.append(minimum4)

                    self.totalMaximum1.append(maximum1)
                    self.totalMaximum2.append(maximum2)
                    self.totalMaximum3.append(maximum3)
                    self.totalMaximum4.append(maximum4)

                self.totalAverage1 = statistics.mean(self.totalAverage1)
                self.totalAverage2 = statistics.mean(self.totalAverage2)
                self.totalAverage3 = statistics.mean(self.totalAverage3)
                self.totalAverage4 = statistics.mean(self.totalAverage4)

                self.totalMinimum1 = min(self.totalMinimum1)
                self.totalMinimum2 = min(self.totalMinimum2)
                self.totalMinimum3 = min(self.totalMinimum3)
                self.totalMinimum4 = min(self.totalMinimum4)

                self.totalMaximum1 = max(self.totalMaximum1)
                self.totalMaximum2 = max(self.totalMaximum2)
                self.totalMaximum3 = max(self.totalMaximum3)
                self.totalMaximum4 = max(self.totalMaximum4)

                self.totalAverage1 = f"{self.totalAverage1:.2f}"
                self.totalAverage2 = f"{self.totalAverage2:.2f}"
                self.totalAverage3 = f"{self.totalAverage3:.2f}"
                self.totalAverage4 = f"{self.totalAverage4:.2f}"

                self.totalMinimum1 = f"{self.totalMinimum1:.2f}"
                self.totalMinimum2 = f"{self.totalMinimum2:.2f}"
                self.totalMinimum3 = f"{self.totalMinimum3:.2f}"
                self.totalMinimum4 = f"{self.totalMinimum4:.2f}"

                self.totalMaximum1 = f"{self.totalMaximum1:.2f}"
                self.totalMaximum2 = f"{self.totalMaximum2:.2f}"
                self.totalMaximum3 = f"{self.totalMaximum3:.2f}"
                self.totalMaximum4 = f"{self.totalMaximum4:.2f}"

                break

            except:
                self.totalAverage1 = "No Data Found"
                self.totalAverage2 = "No Data Found"
                self.totalAverage3 = "No Data Found"
                self.totalAverage4 = "No Data Found"

                self.totalMinimum1 = "No Data Found"
                self.totalMinimum2 = "No Data Found"
                self.totalMinimum3 = "No Data Found"
                self.totalMinimum4 = "No Data Found"

                self.totalMaximum1 = "No Data Found"
                self.totalMaximum2 = "No Data Found"
                self.totalMaximum3 = "No Data Found"
                self.totalMaximum4 = "No Data Found"

        print(f"Selected Total Average: {self.totalAverage1}")
        print(f"Selected Total Average: {self.totalAverage2}")
        print(f"Selected Total Average: {self.totalAverage3}")
        print(f"Selected Total Average: {self.totalAverage4}")

        print(f"Selected Total Minimum: {self.totalMinimum1}")
        print(f"Selected Total Minimum: {self.totalMinimum2}")
        print(f"Selected Total Minimum: {self.totalMinimum3}")
        print(f"Selected Total Minimum: {self.totalMinimum4}")

        print(f"Selected Total Maximum: {self.totalMaximum1}")
        print(f"Selected Total Maximum: {self.totalMaximum2}")
        print(f"Selected Total Maximum: {self.totalMaximum3}")
        print(f"Selected Total Maximum: {self.totalMaximum4}")

#%%
class Tensile():
    tensileData = ""

    rateOfChangeTotalAverage = []
    rateOfChangeTotalMinimum = []
    rateOfChangeTotalMaximum = []

    startForceTotalAverage = []
    startForceTotalMinimum = []
    startForceTotalMaximum = []

    terminatingForceTotalAverage = []
    terminatingForceTotalMinimum = []
    terminatingForceTotalMaximum = []

    readingYear = ""

    fileList = []
    fileFinishedReading = False

    def __init__(self):
        pass
    
    def GettingData(self, itemCode, lotNo):
        self.fileList = []

        if itemCode == "DFB6600600":
            self.fileList = TENSILEData

        # self.rateOfChangeTotalAverage = []
        # self.rateOfChangeTotalMinimum = []
        # self.rateOfChangeTotalMaximum = []

        # self.startForceTotalAverage = []
        # self.startForceTotalMinimum = []
        # self.startForceTotalMaximum = []

        # self.terminatingForceTotalAverage = []
        # self.terminatingForceTotalMinimum = []
        # self.terminatingForceTotalMaximum = []

        #Skipping 4 Rows
        data = self.fileList[0].iloc[4:]

        #Filtering Lot Number Row With Lot Number Input
        tensileLotNoFiltered = data[(data.iloc[:, 3].isin([lotNo]))]

        #Averaging The Rate Of Change Column
        rateOfChangeAverage = round(tensileLotNoFiltered.iloc[:, 6].mean() * 100, 1)
        rateOfChangeMin = round(tensileLotNoFiltered.iloc[:, 6].min() * 100, 1)
        rateOfChangeMax = round(tensileLotNoFiltered.iloc[:, 6].max() * 100, 1)

        self.rateOfChangeTotalAverage = f"{rateOfChangeAverage}%"
        self.rateOfChangeTotalMinimum = f"{rateOfChangeMin}%"
        self.rateOfChangeTotalMaximum = f"{rateOfChangeMax}%"

        #Averaging The Start Force Column
        startForceAverage = tensileLotNoFiltered.iloc[:, 10].mean()
        startForceMin = tensileLotNoFiltered.iloc[:, 10].min()
        startForceMax = tensileLotNoFiltered.iloc[:, 10].max()

        self.startForceTotalAverage = f"{startForceAverage:.1f}"
        self.startForceTotalMinimum = f"{startForceMin:.1f}"
        self.startForceTotalMaximum = f"{startForceMax:.1f}"

        #Averaging The Terminating Column
        terminatingForceAverage = tensileLotNoFiltered.iloc[:, 11].mean()
        terminatingForceMin = tensileLotNoFiltered.iloc[:, 11].min()
        terminationForceMax = tensileLotNoFiltered.iloc[:, 11].max()

        self.terminatingForceTotalAverage = f"{terminatingForceAverage:.1f}"
        self.terminatingForceTotalMinimum = f"{terminatingForceMin:.1f}"
        self.terminatingForceTotalMaximum = f"{terminationForceMax:.1f}"
        
        print(f"RATE OF CHANGE\nAVERAGE: {self.rateOfChangeTotalAverage}\nMINIMUM: {self.rateOfChangeTotalMinimum}\nMAXIMUM: {self.rateOfChangeTotalMaximum}")
        print(f"START FORCE\nAVERAGE: {self.startForceTotalAverage}\nMINIMUM: {self.startForceTotalMinimum}\nMAXIMUM: {self.startForceTotalMaximum}")
        print(f"TERMINATING FORCE\nAVERAGE: {self.terminatingForceTotalAverage}\nMINIMUM: {self.terminatingForceTotalMinimum}\nMAXIMUM: {self.terminatingForceTotalMaximum}")


#%%
# tensile = Tensile()
# tensile.GettingData("DFB6600600", "T000727-02"[:-3])

# %%
