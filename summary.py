"""This script reads an excel spreadsheet and generates a summary sheet.
"""
import numpy as np 
import pandas as pd
import xlwings as xw


class WorkBook:
    """Workbook object to manipulate excel with xlwings."""
    def __init__(self, fn, fragmentation_matrix='fragmentation_matrix.csv',
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
        frag_m = pd.read_csv(fragmentation_matrix, index_col=0)
        self.fragmentation_matrix = frag_m
        self.sheets = wb.sheets
        self.sections = sections
        self.num_sheets = len(self.sheets)
        self.sheet_names = [sht.name for sht in self.sheets]
        print("Found {} sheets with names:\n {}".format(
              self.num_sheets, self.sheet_names))

    def load_fragmentation(self, frag_fn):
        """Load fragmentation database and read relevant values.
        
        Generate frag matrix.

        Parameters
        ----------
        frag_fn : str
            the path to fragmentation file

        Returns
        -------
        None
        """
        pass

    def get_chemicals(self):
        """Get the chemical formula, and multipliers for all sheets."""
        chemicals = input(
            "Input chemicals and multipliers for sheets:\n {}" 
            "\nseparate with comma\n>>>".format(self.sheet_names[1:])
        )
        chemicals = np.array(chemicals.split(',')).reshape(-1, 2)
        self.chemicals = chemicals[:, 0].astype(str)
        self.multipliers = chemicals[:, 1].astype(np.float32)
        
    def get_temp_and_pulse(self):
        """Get temperature and pulse numbers."""
        temp= np.linspace(799, 800, 181)
        pulse = np.ones(temp.shape[0])
        for i in range(pulse.shape[0]):
            pulse[i] = pulse[i] + i * 9
        temp_pulse = np.vstack([temp, pulse]).T
        df_temp_pulse = pd.DataFrame(
            temp_pulse, columns=['temperature', 'pulse']
        )
        return df_temp_pulse

    def _dummy_get_chemicals(self):
        """For test only, skip manual input of chemicals."""
        self.chemicals = ['13Ch4', 'H2O', '13C2H6', '13C2H4',
                          '13CO', '13CO2', 'H2', '13CH4-2']
        self.multipliers = np.ones(len(self.chemicals))

    def get_section0(self):
        """Section 0: sheet names and chemical formula"""
        df_sec0 = pd.DataFrame(np.array([self.chemicals]), columns=self.sheet_names[1:])
        return (self.sections[0], df_sec0)

    def get_section1(self):
        """Section 1, gain correct at gain 7
        
        The values are based on the input sheets.
        """
        df_sec1 = np.empty((181, len(self.chemicals)))
        # dummy coefficients
        cof = np.linspace(0.1, 1, df_sec1.shape[1])
        for i in range(len(self.chemicals)):
            df_sec1[:, i] = self.multipliers[i] * np.array(
                self.sheets[self.sheet_names[i+1]].range('D2', 'D182').value
            )
        self.df_section1 = pd.DataFrame(df_sec1, columns=self.chemicals)
        return (self.sections[1], self.df_section1)

    def get_section2(self):
        """Fragmentation correction."""
        df_sec2 = np.empty(self.df_section1.values.shape)
        for i in range(len(self.chemicals)):
            df_sec2[:, i] = np.dot(
                self.df_section1.values, self.fragmentation_matrix.values[i, :].T
            )
        return (self.sections[2], df_sec2)


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
    df_temp_pulse = wb.get_temp_and_pulse()
    section0 = wb.get_section0()
    section1 = wb.get_section1()
    section2 = wb.get_section2()
    wb.sheets.add('summary', after=wb.sheet_names[-1])
    summary = wb.sheets['summary']

    # write values into summary
    # write temperature and pulse
    summary.range('A2').options(index=False).value = df_temp_pulse
    # write section 0, title: MO, content: chemicals, sheet names
    summary.range('C1').value = section0[0] 
    summary.range('C2').options(index=False).value = section0[1]
    # write  section 1, gain correct at gain 7
    summary.range('K1').value = section1[0]
    summary.range('K2').options(index=False).value = section1[1]
    # write section 2, framentation orrection
    summary.range('S1').value = section2[0]
    summary.range('S2').options(index=False).value = section2[1]


if __name__ == '__main__':
    main()
