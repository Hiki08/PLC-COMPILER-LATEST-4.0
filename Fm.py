#%%
from Imports import *
import DateAndTimeManager
from FilesReader import *

#%%
class fM():
    fmData = ""
    fmItemCode = ""

    totalAverage1 = []
    totalAverage2 = []
    totalAverage3 = []
    totalAverage4 = []
    totalAverage5 = []
    totalAverage6 = []
    totalAverage7 = []

    totalMinimum1 = []
    totalMinimum2 = []
    totalMinimum3 = []
    totalMinimum4 = []
    totalMinimum5 = []
    totalMinimum6 = []
    totalMinimum7 = []
    
    totalMaximum1 = []
    totalMaximum2 = []
    totalMaximum3 = []
    totalMaximum4 = []
    totalMaximum5 = []
    totalMaximum6 = []
    totalMaximum7 = []

    readingYear = ""
    fileFinishedReading = False
    fileList = []

    isValueRetrieve = False

    def __init__(self):
        pass
    
    def GettingData(self, itemCode, lotNumber):
        if itemCode == "FM05000102-00A" or itemCode == "FM05000102-01A":
            self.fileList = FM05000102Data

        for fileNum in range(len(self.fileList)):
            self.totalAverage1 = []
            self.totalAverage2 = []
            self.totalAverage3 = []
            self.totalAverage4 = []
            self.totalAverage5 = []
            self.totalAverage6 = []
            self.totalAverage7 = []

            self.totalMinimum1 = []
            self.totalMinimum2 = []
            self.totalMinimum3 = []
            self.totalMinimum4 = []
            self.totalMinimum5 = []
            self.totalMinimum6 = []
            self.totalMinimum7 = []

            self.totalMaximum1 = []
            self.totalMaximum2 = []
            self.totalMaximum3 = []
            self.totalMaximum4 = []
            self.totalMaximum5 = []
            self.totalMaximum6 = []
            self.totalMaximum7 = []

            print(f"READING FILE {fileNum}")

            #Getting The Row, Column Location Of HIBLOW
            findSupplier = [(index, column) for index, row in self.fileList[fileNum].iterrows() for column, value in row.items() if value == "SUPPLIER"]
            supplierRow = [index for index, _ in findSupplier]
            supplierColumn = [column for _, column in findSupplier]

            print("Row indices:", supplierRow)
            print("Column names:", supplierColumn)

            # Get the Neighboring Data Of SUPPLIER
            supplierFiltered = self.fileList[fileNum].iloc[max(0, supplierRow[0] - 3):min(len(self.fileList[fileNum]), supplierRow[0] + 10), self.fileList[fileNum].columns.get_loc(supplierColumn[0]):self.fileList[fileNum].columns.get_loc(supplierColumn[0]) + 999999]
            # return supplierFiltered

            try:
                #Getting The Row, Column Location Of Lot Number
                findLotNumber = [(index, column) for index, row in supplierFiltered.iterrows() for column, value in row.items() if value == lotNumber]
                lotNumberRow = [index for index, _ in findLotNumber]
                lotNumberColumn = [column for _, column in findLotNumber]

                print("Row indices:", lotNumberRow)
                print("Column names:", lotNumberColumn)

                for a in range(0, len(lotNumberColumn)):
                    # Get The Neighboring Data of Lot Number
                    inspectionData = self.fileList[fileNum].iloc[max(0, lotNumberRow[a]):min(len(self.fileList[fileNum]), lotNumberRow[a] + 13), self.fileList[fileNum].columns.get_loc(lotNumberColumn[a]):self.fileList[fileNum].columns.get_loc(lotNumberColumn[a]) + 5]

                    #CHECKING THE ITEM CODE
                    if itemCode == "FM05000102-00A" or itemCode == "FM05000102-01A":
                        average1 = inspectionData.iloc[3].mean()
                        average2 = inspectionData.iloc[4].mean()
                        average3 = inspectionData.iloc[5].mean()
                        average4 = inspectionData.iloc[6].mean()
                        average5 = inspectionData.iloc[7].mean()
                        average6 = inspectionData.iloc[8].mean()
                        average7 = inspectionData.iloc[9].mean()

                        minimum1 = inspectionData.iloc[3].min()
                        minimum2 = inspectionData.iloc[4].min()
                        minimum3 = inspectionData.iloc[5].min()
                        minimum4 = inspectionData.iloc[6].min()
                        minimum5 = inspectionData.iloc[7].min()
                        minimum6 = inspectionData.iloc[8].min()
                        minimum7 = inspectionData.iloc[9].min()

                        maximum1 = inspectionData.iloc[3].max()
                        maximum2 = inspectionData.iloc[4].max()
                        maximum3 = inspectionData.iloc[5].max()
                        maximum4 = inspectionData.iloc[6].max()
                        maximum5 = inspectionData.iloc[7].max()
                        maximum6 = inspectionData.iloc[8].max()
                        maximum7 = inspectionData.iloc[9].max()

                        self.totalAverage1.append(average1)
                        self.totalAverage2.append(average2)
                        self.totalAverage3.append(average3)
                        self.totalAverage4.append(average4)
                        self.totalAverage5.append(average5)
                        self.totalAverage6.append(average6)
                        self.totalAverage7.append(average7)

                        self.totalMinimum1.append(minimum1)
                        self.totalMinimum2.append(minimum2)
                        self.totalMinimum3.append(minimum3)
                        self.totalMinimum4.append(minimum4)
                        self.totalMinimum5.append(minimum5)
                        self.totalMinimum6.append(minimum6)
                        self.totalMinimum7.append(minimum7)

                        self.totalMaximum1.append(maximum1)
                        self.totalMaximum2.append(maximum2)
                        self.totalMaximum3.append(maximum3)
                        self.totalMaximum4.append(maximum4)
                        self.totalMaximum5.append(maximum5)
                        self.totalMaximum6.append(maximum6)
                        self.totalMaximum7.append(maximum7)

                    elif itemCode == "FM03500100-01":
                        # IN PROGRESS
                        pass

                if itemCode == "FM05000102-00A" or itemCode == "FM05000102-01A":
                    self.totalAverage1 = statistics.mean(self.totalAverage1)
                    self.totalAverage2 = statistics.mean(self.totalAverage2)
                    self.totalAverage3 = statistics.mean(self.totalAverage3)
                    self.totalAverage4 = statistics.mean(self.totalAverage4)
                    self.totalAverage5 = statistics.mean(self.totalAverage5)
                    self.totalAverage6 = statistics.mean(self.totalAverage6)
                    self.totalAverage7 = statistics.mean(self.totalAverage7)

                    self.totalMinimum1 = min(self.totalMinimum1)
                    self.totalMinimum2 = min(self.totalMinimum2)
                    self.totalMinimum3 = min(self.totalMinimum3)
                    self.totalMinimum4 = min(self.totalMinimum4)
                    self.totalMinimum5 = min(self.totalMinimum5)
                    self.totalMinimum6 = min(self.totalMinimum6)
                    self.totalMinimum7 = min(self.totalMinimum7)

                    self.totalMaximum1 = max(self.totalMaximum1)
                    self.totalMaximum2 = max(self.totalMaximum2)
                    self.totalMaximum3 = max(self.totalMaximum3)
                    self.totalMaximum4 = max(self.totalMaximum4)
                    self.totalMaximum5 = max(self.totalMaximum5)
                    self.totalMaximum6 = max(self.totalMaximum6)
                    self.totalMaximum7 = max(self.totalMaximum7)

                    self.totalAverage1 = f"{self.totalAverage1:.2f}"
                    self.totalAverage2 = f"{self.totalAverage2:.2f}"
                    self.totalAverage3 = f"{self.totalAverage3:.2f}"
                    self.totalAverage4 = f"{self.totalAverage4:.2f}"
                    self.totalAverage5 = f"{self.totalAverage5:.2f}"
                    self.totalAverage6 = f"{self.totalAverage6:.2f}"
                    self.totalAverage7 = f"{self.totalAverage7:.2f}"
                    
                    self.totalMinimum1 = f"{self.totalMinimum1:.2f}"
                    self.totalMinimum2 = f"{self.totalMinimum2:.2f}"
                    self.totalMinimum3 = f"{self.totalMinimum3:.2f}"
                    self.totalMinimum4 = f"{self.totalMinimum4:.2f}"
                    self.totalMinimum5 = f"{self.totalMinimum5:.2f}"
                    self.totalMinimum6 = f"{self.totalMinimum6:.2f}"
                    self.totalMinimum7 = f"{self.totalMinimum7:.2f}"

                    self.totalMaximum1 = f"{self.totalMaximum1:.2f}"
                    self.totalMaximum2 = f"{self.totalMaximum2:.2f}"
                    self.totalMaximum3 = f"{self.totalMaximum3:.2f}"
                    self.totalMaximum4 = f"{self.totalMaximum4:.2f}"
                    self.totalMaximum5 = f"{self.totalMaximum5:.2f}"
                    self.totalMaximum6 = f"{self.totalMaximum6:.2f}"
                    self.totalMaximum7 = f"{self.totalMaximum7:.2f}"

                    break

                elif itemCode == "FM03500100-01":
                    # IN PROGRESS
                    pass
                
            except:
                self.totalAverage1 = "No Data Found"
                self.totalAverage2 = "No Data Found"
                self.totalAverage3 = "No Data Found"
                self.totalAverage4 = "No Data Found"
                self.totalAverage5 = "No Data Found"
                self.totalAverage6 = "No Data Found"
                self.totalAverage7 = "No Data Found"

                self.totalMinimum1 = "No Data Found"
                self.totalMinimum2 = "No Data Found"
                self.totalMinimum3 = "No Data Found"
                self.totalMinimum4 = "No Data Found"
                self.totalMinimum5 = "No Data Found"
                self.totalMinimum6 = "No Data Found"
                self.totalMinimum7 = "No Data Found"

                self.totalMaximum1 = "No Data Found"
                self.totalMaximum2 = "No Data Found"
                self.totalMaximum3 = "No Data Found"
                self.totalMaximum4 = "No Data Found"
                self.totalMaximum5 = "No Data Found"
                self.totalMaximum6 = "No Data Found"
                self.totalMaximum7 = "No Data Found"

            print(f"Total Average: {self.totalAverage1}")
            print(f"Total Average: {self.totalAverage2}")
            print(f"Total Average: {self.totalAverage3}")
            print(f"Total Average: {self.totalAverage4}")
            print(f"Total Average: {self.totalAverage5}")
            print(f"Total Average: {self.totalAverage6}")
            print(f"Total Average: {self.totalAverage7}")

            print(f"Total Minimum: {self.totalMinimum1}")
            print(f"Total Minimum: {self.totalMinimum2}")
            print(f"Total Minimum: {self.totalMinimum3}")
            print(f"Total Minimum: {self.totalMinimum4}")
            print(f"Total Minimum: {self.totalMinimum5}")
            print(f"Total Minimum: {self.totalMinimum6}")
            print(f"Total Minimum: {self.totalMinimum7}")

            print(f"Total Maximum: {self.totalMaximum1}")
            print(f"Total Maximum: {self.totalMaximum2}")
            print(f"Total Maximum: {self.totalMaximum3}")
            print(f"Total Maximum: {self.totalMaximum4}")
            print(f"Total Maximum: {self.totalMaximum5}")
            print(f"Total Maximum: {self.totalMaximum6}")
            print(f"Total Maximum: {self.totalMaximum7}")

        print(f"Selected Total Average: {self.totalAverage1}")
        print(f"Selected Total Average: {self.totalAverage2}")
        print(f"Selected Total Average: {self.totalAverage3}")
        print(f"Selected Total Average: {self.totalAverage4}")
        print(f"Selected Total Average: {self.totalAverage5}")
        print(f"Selected Total Average: {self.totalAverage6}")
        print(f"Selected Total Average: {self.totalAverage7}")

        print(f"Selected Total Minimum: {self.totalMinimum1}")
        print(f"Selected Total Minimum: {self.totalMinimum2}")
        print(f"Selected Total Minimum: {self.totalMinimum3}")
        print(f"Selected Total Minimum: {self.totalMinimum4}")
        print(f"Selected Total Minimum: {self.totalMinimum5}")
        print(f"Selected Total Minimum: {self.totalMinimum6}")
        print(f"Selected Total Minimum: {self.totalMinimum7}")

        print(f"Selected Total Maximum: {self.totalMaximum1}")
        print(f"Selected Total Maximum: {self.totalMaximum2}")
        print(f"Selected Total Maximum: {self.totalMaximum3}")
        print(f"Selected Total Maximum: {self.totalMaximum4}")
        print(f"Selected Total Maximum: {self.totalMaximum5}")
        print(f"Selected Total Maximum: {self.totalMaximum6}")
        print(f"Selected Total Maximum: {self.totalMaximum7}")