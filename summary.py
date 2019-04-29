"""This script reads an excel spreadsheet and generates a summary sheet.
"""
import numpy as np 
import pandas as pd
from matplotlib import rcParams
rcParams['backend'] = 'agg'
import matplotlib.pyplot as plt
import xlwings as xw


class WorkBook:
    """Workbook object to manipulate excel with xlwings."""
    def __init__(self, fn, fragmentation_matrix='fragmentation_matrix.csv',
                 gain_setting='gain_setting.xlsx',
                 sections=['M0', 'gain correct at gain 7',
                           'fragmentation correction',
                           'relative',
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
        # delete summary sheet if already there
        if self.wb.sheets[-1].name == 'summary':
            print("Delete existing summary sheet " 
                  "\nmake a new one after analysis")
            self.wb.sheets['summary'].delete()
        self.sheets = self.wb.sheets
        self.sections = sections
        self.num_sheets = len(self.sheets)
        self.sheet_names = [sht.name for sht in self.sheets]
        # for compatibility, keep chemicals attributes
        self.chemicals = self.sheet_names
        print("Found {} sheets with names:\n {}".format(
              self.num_sheets, self.sheet_names))
        # get gain settings
        print("Use gain setting file: {}".format(gain_setting))
        gain_df = pd.read_excel(gain_setting)
        self.multipliers = gain_df.Factor.values

    #def _get_gain_multipliers(self):
    #    """Load multipliers for gain.
    #    
    #    Read it from the excel file
    #    """
    #    # this is a temporary dummy multiplers
    #    multipliers = self.gain_df.Factor.values
    #    return multipliers

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

    #def get_chemicals(self):
    #    """Get the chemical formula, and multipliers for all sheets."""
    #    chemicals = input(
    #        "Input chemicals and multipliers for sheets:\n {}" 
    #        "\nseparate with comma\n>>>".format(self.sheet_names)
    #    )
    #    chemicals = np.array(chemicals.split(',')).reshape(-1, 2)
    #    self.chemicals = chemicals[:, 0].astype(str)
    #    self.multipliers = chemicals[:, 1].astype(np.float32)
        
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
        self.df_temp_pulse = df_temp_pulse
        return df_temp_pulse

    #def dummy_get_chemicals(self):
    #    """For test only, skip manual input of chemicals."""
    #    self.chemicals = ['Ar', '13Ch4', 'H2O', '13C2H6', '13C2H4',
    #                      '13CO', '13CO2', 'H2', '13CH4-2']
    #    self.multipliers = np.ones(len(self.chemicals))

    def get_section0(self):
        """Section 0: sheet names and chemical formula"""
        # read M0 of each sheet
        vals = np.zeros((181, len(self.chemicals)))
        for i in range(len(self.chemicals)):
            vals[:, i] = np.array(
                self.sheets[self.sheet_names[i]].range('A2', 'A182').value
            )
        # use sheetnames and chemicals as columns
        cols = self.chemicals
        self.df_section0 = pd.DataFrame(vals, columns=cols)
        return (self.sections[0], self.df_section0)

    def get_section1(self):
        """Section 1, gain correct at gain 7
        
        The values are based on section0.
        """
        # dummy coefficients
        #cof = np.linspace(0.1, 1, self.df_section0.shape[1])
        vals = self.df_section0.values * self.multipliers
        self.df_section1 = pd.DataFrame(vals, columns=self.chemicals)

        return (self.sections[1], self.df_section1)

    def get_section2(self):
        """Fragmentation correction."""
        vals = np.empty(self.df_section1.values.shape)
        for i in range(len(self.chemicals)):
            vals[:, i] = np.dot(
                self.df_section1.values, 
                self.fragmentation_matrix.values[i, :].T
            )
        self.df_section2 = pd.DataFrame(vals, columns=self.chemicals)
        return (self.sections[2], self.df_section2)

    def get_section3(self):
        """Relative.
        
        Fragmentation divided by inert.
        """
        # get inert values and convert to column vector
        inert = self.df_section2.iloc[:, 0].values[:, np.newaxis]
        # divide each col in section 1 with inert 
        vals = self.df_section2.values / inert
        self.df_section3 = pd.DataFrame(vals, columns=self.chemicals)
        return (self.sections[3], self.df_section3)

    def plot_section_relative(self, size=(5, 3)):
        """Plot section relative against pulse number."""
        figures = []
        xx = self.df_temp_pulse.values[:, 1]
        yys = self.df_section3.values
        for i, chem in enumerate(self.chemicals):
            fig = plt.figure(figsize=size)
            plt.plot(xx, yys[:, i])
            plt.title(chem)
            figures.append(fig)
        return figures


def main():
    import sys
    import warnings
    if len(sys.argv) == 1:
        warnings.warn("No excel found")
        fn = input("input excel filename\n>>>")
    else:
        fn = sys.argv[1]
    wb = WorkBook(fn)
    #wb.dummy_get_chemicals()
    print("Process data")
    df_temp_pulse = wb.get_temp_and_pulse()
    section0 = wb.get_section0()  # M0
    section1 = wb.get_section1()  # gain correction
    section2 = wb.get_section2()  # fragmentation
    section3 = wb.get_section3()  # relative
    figures = wb.plot_section_relative()  # plot section relative
    print("Create sheet 'summary'")
    wb.sheets.add('summary', after=wb.sheet_names[-1])
    summary = wb.sheets['summary']

    print("Write data to summary")
    # write temperature and pulse
    summary.range('A2').options(index=False).value = df_temp_pulse
    # write section 0, title: MO, content: chemicals, sheet names
    summary.range('C1').value = section0[0] 
    summary.range('C2').options(index=False).value = section0[1]
    # write  section 1, gain correct at gain 7
    summary.range('L1').value = section1[0]
    summary.range('L2').options(index=False).value = section1[1]
    # write section 2, framentation orrection
    summary.range('U1').value = section2[0]
    summary.range('U2').options(index=False).value = section2[1]
    # write section 3, relative
    summary.range('AD1').value = section3[0]
    summary.range('AD2').options(index=False).value = section3[1]
    # put plots of relative at the end
    print("Plot relative section")
    for i, fig in enumerate(figures):
        left = 2500
        top = i * 300
        summary.pictures.add(fig, left=left, top=top)

    # autofit width
    summary.autofit('c')
    wb.wb.save()
    print("Complete, save excel")

if __name__ == '__main__':
    main()
