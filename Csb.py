#%%
from Imports import *
import DateAndTimeManager
from FilesReader import *

#%%
class cSB():
    csbData = ""
    csbItemCode = ""

    totalAverage1 = []
    
    totalMinimum1 = []

    totalMaximum1 = []
    
    readingYear = ""
    fileFinishedReading = False
    fileList = []

    def __init__(self):
        pass
    
    def GettingData(self, itemCode, lotNumber):
        if itemCode == "CSB6400802":
            self.fileList = CSB6400802Data

        for fileNum in range(len(self.fileList)):
            self.totalAverage1 = []

            self.totalMinimum1 = []

            self.totalMaximum1 = []

            print(f"READING FILE {fileNum}")

            #Getting The Row, Column Location Of SUPPLIER
            findSupplier = [(index, column) for index, row in self.fileList[fileNum].iterrows() for column, value in row.items() if value == "SUPPLIER"]
            supplierRow = [index for index, _ in findSupplier]
            supplierColumn = [column for _, column in findSupplier]

            print("Row indices:", supplierRow)
            print("Column names:", supplierColumn)

            # Get the Neighboring Data Of Hiblow
            supplierFiltered = self.fileList[fileNum].iloc[max(0, supplierRow[0] - 3):min(len(self.fileList[fileNum]), supplierRow[0] + 10), self.fileList[fileNum].columns.get_loc(supplierColumn[0]):self.fileList[fileNum].columns.get_loc(supplierColumn[0]) + 999999]

            try:
                #Getting The Row, Column Location Of Lot Number
                findLotNumber = [(index, column) for index, row in supplierFiltered.iterrows() for column, value in row.items() if value == lotNumber]
                lotNumberRow = [index for index, _ in findLotNumber]
                lotNumberColumn = [column for _, column in findLotNumber]

                print("Row indices:", lotNumberRow)
                print("Column names:", lotNumberColumn)

                for a in range(0, len(lotNumberColumn)):
                    # Get The Neighboring Data of Lot Number
                    inspectionData = self.fileList[fileNum].iloc[max(0, lotNumberRow[a]):min(len(self.fileList[fileNum]), lotNumberRow[a] + 10), self.fileList[fileNum].columns.get_loc(lotNumberColumn[a]):self.fileList[fileNum].columns.get_loc(lotNumberColumn[a]) + 5]

                    #CHECKING THE ITEM CODE
                    if itemCode == "CSB6400802":
                        average1 = inspectionData.iloc[3].mean()

                        minimum1 = inspectionData.iloc[3].min()

                        maximum1 = inspectionData.iloc[3].max()

                        self.totalAverage1.append(average1)

                        self.totalMinimum1.append(minimum1)

                        self.totalMaximum1.append(maximum1)
                
                if itemCode == "CSB6400802":
                    self.totalAverage1 = statistics.mean(self.totalAverage1)

                    self.totalMinimum1 = min(self.totalMinimum1)

                    self.totalMaximum1 = max(self.totalMaximum1)

                    self.totalAverage1 = f"{self.totalAverage1:.2f}"

                    self.totalMinimum1 = f"{self.totalMinimum1:.2f}"

                    self.totalMaximum1 = f"{self.totalMaximum1:.2f}"

                    break

            except:
                self.totalAverage1 = "No Data Found"

                self.totalMinimum1 = "No Data Found"

                self.totalMaximum1 = "No Data Found"

        print(f"Selected Total Average: {self.totalAverage1}")

        print(f"Selected Total Minimum: {self.totalMinimum1}")

        print(f"Selected Total Maximum: {self.totalMaximum1}")