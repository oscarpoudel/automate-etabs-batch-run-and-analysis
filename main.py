import os
import sys
import comtypes.client
import pandas as pd
from list_dir import ListDir
from export_tables_to_xls import Runapp
from xls_save import saveExcel
from pywinauto.application import Application
from time import sleep
from pywinauto import keyboard

#Input the model list path here
Search_path=r"D:\P_B_Docu\research_articles\inclined_column\final_files\Etabs Models"
extension=r".EDB"
dir=ListDir(Search_path,extension)
dir.search()


class mainClass:
   
    def EtabsModel(self):

        AttachToInstance = False
        SpecifyPath = False
        helper = comtypes.client.CreateObject('ETABSv1.Helper')
        helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)
        if AttachToInstance:
            # attach to a running instance of ETABS
            try:
                # get the active ETABS object
                myETABSObject = helper.GetObject(
                    "CSI.ETABS.API.ETABSObject")
            except (OSError, comtypes.COMError):
                print(
                    "No running instance of the program found or failed to attach.")
                sys.exit(-1)
        else:

            try:
                # create an instance of the ETABS object from the latest installed ETABS
                myETABSObject = helper.CreateObjectProgID(
                    "CSI.ETABS.API.ETABSObject")
            except (OSError, comtypes.COMError):
                print("Cannot start a new instance of the program.")
                sys.exit(-1)
        # Start the application
        myETABSObject.ApplicationStart()
        # Create a sapmodel
        SapModel = myETABSObject.SapModel
        return SapModel

    def Analyze_file(self, ModelPath, EtabsModel):
        # Open the model files
        SapModel = EtabsModel
        currentModel = SapModel.File.OpenFile(ModelPath)
        # Unlock the model
        SapModel.SetModelIsLocked(0)
        # Run THe Model
        currentModel = SapModel.Analyze.RunAnalysis()
        print('______________Analysis Completed now exporting table__________')
       
        # Go to export table          
        run = Runapp()
        run.runit(ModelPath)

        # sav = saveExcel()
        SapModel.SetModelIsLocked(0)
        SapModel.file.Save()
        SapModel=None
        # sav.runit(ModelPath)

       

    def open_etabs(self):
        etabs_path = r"C:\Program Files\Computers and Structures\ETABS 19\ETABS.exe"
        app = Application(backend='uia').start(etabs_path)

    def open_file(self,path):
        app = Application(backend='uia').connect(
                    title_re=".*ETABS*", timeout=5000)
        topWindow = app.top_window().wait('ready',timeout = 500)
        # try:
        #     topWindow.menu_select('File->Open...')
        # except:
        #     pass
        topWindow.type_keys('^o')
        top_win2= app.top_window()
        # top_win2.print_control_identifiers()
        top_win2.FileNameEdit.set_edit_text(path)
        
        sleep(5)
        keyboard.send_keys("{ENTER}")
        sleep(5)
        try:
             app.top_window().wait('ready',timeout=5000).menu_select('Analyze -> Run Analysis') 
        except:
            pass
        run = Runapp()
        run.runit(path)
        # sav = saveExcel()
        # sav.runit(path)


    def open_dirs(self):
        # self.open_etabs()
        model = self.EtabsModel()
        f = open('edb_list.txt', 'r')
        fil=[]
        for i in f:
            fil.append(i.rstrip('\n'))
        
        i=1
        
        for path in fil:
            print(path)
            # path = (f.readline()).rstrip('\n')
            direc = os.path.dirname(path)
            for fname in os.listdir(direc):
                if fname.endswith('.xlsx'):
                    print('Excel already exists for')
                    loop=False
                    break
                else:
                    loop=True
            if loop:
                    
                    self.Analyze_file(path, model)
                    # self.open_file(path)
                    print(f'[${i}]---done for file: ${path}--------')
                    print('XXXXXXXXXXXXxxxxxxxxxXXXXXXXXXXXXDONEXXXXXxxxxxxxxxXXXXXX\n\n\n')
                    i=i+1
        


if __name__ == '__main__':
    run = mainClass()
    run.open_dirs()
