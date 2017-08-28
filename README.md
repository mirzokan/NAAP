# Usage
* Download the Excel workbook file
* Read the detailed instructions in the file

# Description
This Excel app analyzes NECEEM electropherograms and extracts <i>K</i><sub>d</sub>, <i>k</i><sub>on</sub> and <i>k</i><sub>off</sub> values in batch. It employs the peak deconvolution method developed by [Cherney et al](http://pubs.acs.org/doi/abs/10.1021/ac2027113). This version is developed for analysis of data obtained by a P/ACE MDQ series of Capillary Electrophoresis instruments (Sciex), and exported by 32 Karat software, version 7. Requires Microsoft Excel version 2007 or later. 

The general algorithm of the program:

* All data files within a user-defined folder are imported into the workbook
* The baseline is linearised within the user-defined working region of the signal
* User-defined background regions and a signal-to-noise constant are used to identify non-convoluted edges of the complex and free-ligand peaks
* The data is smoothed using  Savitzky-Golay method
* A first derivative of the data is used to approximate maxima of the two peaks. The peak maxima values are further refined by fitting to a polynomial equation. 
* Assuming peak symmetry, convoluted edges of the two peaks are resolved and areas under the three regions of the NECEEM electropherograms are calculated. This step uses the original non-smoothed data
* The signal areas are used to calculate Kd, kon and koff.
