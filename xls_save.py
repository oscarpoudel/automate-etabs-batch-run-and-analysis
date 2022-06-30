from time import sleep
from pywinauto import Application
from pywinauto.keyboard import send_keys
import os
import shutil


class saveExcel:
    def runit(self, path='./'):
        app = Application(backend='uia').connect(
            title_re=".*ETABS*", timeout=5000)
        display_tabl=app.top_window()
        display_tabl.set_focus()

        try:
            print('Enterred inside closing assenbling')
            table_windo = display_tabl.child_window(auto_id="fReportDB")
            table_windo.close()
            display_tabl.menu_select('Analyze -> Unlock Model')
            sleep(2)
            display_tabl[u'Ok'].click()
            sleep(2)
            display_tabl.wait('ready').type_keys('^s')

            print('Finished inside closing assenbling')
        except:
            print('Assembled not found')
        excel = Application(backend='uia').connect(title_re='.*- Excel*',timeout=50000)
        excel_win = excel.window(title_re='.*- Excel*')
        print('waiting for the excel file')
        
        try:
            excel_win.wait("exists ready",timeout=50000).type_keys('{F12}')
        except:
            pass
        directory = os.path.dirname(path)
        print(directory)
        save_window = excel['Dialog']
        save_window["File name:Edit"].type_keys('D:\\analysis_results', with_spaces=True)
        save_window[u'Save'].click()
        sleep(15)
        excel_win.type_keys("%{F4}")
        
        try:
            excel_win.close()
        except:
            pass
        sleep(5)
        shutil.move('D:\\analysis_results.xlsx', directory)
        print('____________________________Saving Process Completed_________________')


if __name__ == '__main__':
    l = saveExcel()

    l.runit()
