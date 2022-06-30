import sys
import os

class ListDir:
    def __init__(self,path,extension):
        self.path = path
        self.extension= extension

    def search(self):
            
            org_stdout = sys.stdout
            file ='edb_list.txt'
            sys.stdout=open(file,'w')

            for root,dirs,files in os.walk(self.path):
                for file in files:
                    if file.endswith(self.extension):
                        
                            print(os.path.join(root,file))
            sys.stdout=org_stdout   
            print('listing directory done')        

