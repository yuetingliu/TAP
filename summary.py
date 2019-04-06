"""This script reads an excel spreadsheet and generates a summary sheet.
"""
import numpy as np 
import pandas as pd
import xlwings as xw

xlsx = 'sample.xlsx'

wb = xw.Book(xlsx)
sht = wb.sheets['inert']

print("sheets", wb.sheets)
print("sheet inert", sht.range('A1').value)

class WorkBook:
    """Workbook object to manipulate excel with xlwings."""
    def __init__(self, fn,
                 sections=['M0', 'gain correct at gain 7',
                           'fragmentation correction',
                           'sensitivity correction']):
        """Read an excel file and store the status.
        
        Parameters
        ----------
        fn : str
            the filename of execl

        Returns
        -------
        None
        """
        self.wb = xw.Book(fn)
        self.sheets = wb.sheets
        self.sections = sections
        self.num_sheets = len(self.sheets)
        self.sheet_names = [sht.name for sht in self.sheets]
        print("Found {} sheets with names:\n {}".format(
              self.num_sheets, self.sheet_names))

    def get_chemicals(self):
        """Get the chemical formula for all sheets."""
        chemicals = input("Input chemicals for sheets:\n {}" 
                          "\nseparate with comma\n>>>".format(self.sheet_names))
        self.chemicals = chemicals.split(',')
        
    def get_temp_pulse(self):
        """Get temperature and pulse numbers."""
        temperature = np.linspace(799, 800, 181)
        pulse = np.ones(temperature.shape[0])
        for i in range(pulse.shape[0]):
            pulse[i] = pulse[i] + i * 9
        self.temperature = temperature
        self.pulse = pulse

    def _dummy_get_chemicals(self):
        """For test only, skip manual input of chemicals."""
        self.chemicals = ['13Ch4', 'H2O', '13C2H6', '13C2H4',
                          '13CO', '13CO2', 'H2', '13CH4-2']

    def get_section0(self):
        """Section 0: sheet names and chemical formula"""
        values = [self.chemicals, self.sheet_names[1:]]
        df_sec1 = pd.DataFrame(values)
        return (self.sections[0], df_sec1)

    def get_section1(self):
        """Section 1
        
        The values are based on the input sheets.
        """
        df_sec2 = np.empty((181, self.num_sheets - 1))
        # dummy coefficients
        cof = np.linspace(0.1, 1, df_sec1.shape[1])
        for i in range(len(self.chemicals)):
            df_sec1[:, i] = cof[i] * np.array(
                self.sheets[self.sheet_names[i+1]].range('D2', 'D182').value
            )
        self.df_section1 = pd.DataFrame(df_sec1, columns=self.chemicals)


def main():
    import sys
    import warnings
    if len(sys.argv) == 1:
        warnings.warn("No excel found")
        fn = input("input excel filename\n>>>")
    else:
        fn = sys.argv[1]
    wb = WorkBook(fn)
    wb.get_chemicals()
    wb.get_sec1_values()


if __name__ == '__main__':
    main()
