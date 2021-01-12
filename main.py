
from shutil import rmtree
import os 
from time import sleep
from datetime import datetime

from modules.arhivete import arhfolder
from modules.dataReader import dataReader
from modules.docsCreator import docsCreator

class docCreators():

    finpyt = 'DataList.xlsx'
    tinpyt = 'Temp.docx'
    filenamecol = 'val_1'
    
    def get(self, zip=False):
        # Определяем адрес корневой директории и базовые файлы
        path = os.path.dirname(os.path.abspath(__file__))
        inputfile = os.path.join(path, 'inputFiles', self.finpyt)
        tempfile = os.path.join(path, 'templateFiles', self.tinpyt)
        resf = 'Results'
        # Наименования столбика с именем файла
        # Считываем входные данные
        data = dataReader(inputfile)[1]

        # Создаем папку для файлов

        try:
            os.mkdir(os.path.join(path, resf))
        except:
            pass

        # Создаем название сессии
        for i in range(0, 100):

            try:
                session = datetime.now().strftime("%Y%m%d%H%M%S")+str(i)
                resdir = os.path.join(path, resf, session)
                os.mkdir(resdir)
                break 
            except:
                pass

        
        # Создаем папку с массивом документов
        for i in data:
            resfile = os.path.join(resdir, i[self.filenamecol]+'.docx')
            docsCreator(tempfile, resfile, i)

        if zip:
            # Создаем архив с массивом документов


            arhfolder(os.path.join(path, resf),
                    session, 
                    path)


            # Удаляем папку с массивом документов
            rmtree(resdir, ignore_errors=False, onerror=None)


docCreators().get()