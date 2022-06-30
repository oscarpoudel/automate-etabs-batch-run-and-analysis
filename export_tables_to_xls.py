from pywinauto.application import Application
import sys
from pywinauto.keyboard import send_keys
from time import sleep
from xls_save import saveExcel
from pywinauto.timings import WaitUntilPasses
import shutil
import os


class Runapp:
    def runit(self,path='D:\\gis\\tes.txt'):
        app = Application(backend='uia').connect(
            title_re=".*ETABS*", timeout=5000)
        print('connected for exporting and saving table')
        # print(app.windows())
        sleep(10)
        main_window = app.window(best_match='ETABS Ultimate 19.1.0',auto_id="mdiMainForm", control_type="Window");
        main_window.wait('ready',timeout=50000)
        try:
            # sleep(10)
            main_window.set_focus()
            sleep(5)
            #main_window.menu_select('Display->Show Tables...')
            # main_window.menu_select('File -> Export -> ETABS Database Tables to Excel..')
            main_window.wait('ready').type_keys('^i')

            print('table found proceeding')

        except:
            print('table not show')

        display_tables = main_window.child_window(title="Choose Tables for Export to Excel", auto_id="DBTableForm", control_type="Window")
        display_tables.wait('ready',timeout=300)
        display_tables.set_focus()
        sleep(2)
        okbtn=display_tables.child_window(title="OK", auto_id="cmdOK", control_type="Button")

        if okbtn.is_enabled() == False:
            display_tables.child_window(title="Joint Output", control_type="TreeItem").click_input()

            display_tables.child_window(title="Structure Output", control_type="TreeItem").click_input()
        okbtn.click()
        # display_tables.print_control_identifiers()

        sleep(5)

        '''
        table_windows = main_window.child_window(title='Please Wait While Table Display Initializes', auto_id="fReportDB")
        
        # table_windows.wait('exists enabled visible ready', timeout=50000)
        table_windows.print_control_identifiers()
        table_windows.wait_not('exists', timeout=500)
        print('passed on from there')
        main_window.child_window(title='Assembled Joint Masses').menu_select(
                'File->Export All Tables->To Excel')  
        del app  


        '''
        saveasWind = main_window.child_window(title="Save As")
        saveasWind.wait('ready')
        saveasWind.FileNameEdit.set_edit_text("D:\\analysis_results")
        saveasWind.child_window(title="Save", auto_id="1", control_type="Button").click()
        '''
        try:
            os.remove(r'D:\P_B_Docu\research_articles\inclined_column\final_files\python_for_etabs\etabs_files\analysis_results.xlsx')
        except:
            pass'''
        directory = os.path.dirname(path)
        print('<<Waiting for exporting process>>')
        while not os.path.exists('D:\\analysis_results.xlsx'):
            sleep(3)
        sleep(15)
        print(os.listdir('D:\\'))
        shutil.move("D:\\analysis_results.xlsx", directory)
        print('______________________Exporting to Excel Completed____________________')
        # ls = saveExcel()
        # path = r'D:\etabs_files\Residential1.EDB'
        # ls.runit()


if __name__ == '__main__':
    ap = Runapp()

    ap.runit()
