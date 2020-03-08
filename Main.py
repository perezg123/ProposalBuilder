from pyforms.basewidget import BaseWidget
from pyforms.controls   import ControlFile
from pyforms.controls   import ControlText
from pyforms.controls   import ControlSlider
from pyforms.controls   import ControlPlayer
from pyforms.controls   import ControlButton
from pyforms.controls   import ControlCombo
import openpyxl

class ProposalBuilder(BaseWidget):

    def __init__(self, *args, **kwargs):
        super().__init__('Proposal Builder')

        #Definition of the forms fields
        self._inputfile  = ControlFile('Price Sheet')
        self._sku = ControlCombo('Item SKU')
        self._outputfile = ControlText('Proposal File')


        #Define the function that will be called when a file is selected
        self._inputfile.changed_event = self.__input_file_selection_event
        #Define the event that will be called when the run button is processed
#        self._runbutton.value = self.run_event
        #Define the event called before showing the image in the player
#        self._player.process_frame_event = self.__process_frame

        #Define the organization of the Form Controls
        self._formset = [
            ('_inputfile', ' ', '_outputfile', '_sku'),
        ]


    def __input_file_selection_event(self):
        """
        When the inputfile congtrol is selected
        """
        part_data = {}
        self._inputfile.value = self._inputfile.value
        wb = openpyxl.load_workbook('example.xlsx')
        sheet = wb.get_sheet_by_name('FortiGate')
        for row in range(2, sheet.max_row + 1):
            if (sheet['A' + str(row)].value == "UNIT") and (sheet['B'+ str(row)].value == "SKU"):
                header_row = True
                continue
            else:
                part_data =

    def __process_frame(self, frame):
        """
        Do some processing to the frame and return the result frame
        """
        return frame

    def run_event(self):
        """
        After setting the best parameters run the full algorithm
        """
        print("The function was executed", self._inputfile.value)


if __name__ == '__main__':

    from pyforms import start_app
    start_app(ProposalBuilder, geometry=(200, 200, 300, 300))