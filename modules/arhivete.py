import zipfile
import os

def arhfolder(zippath, aname, sfold):

    resfile = os.path.join(zippath, aname+'.zip')
    zipfd = zipfile.ZipFile(resfile, 'w', zipfile.ZIP_DEFLATED)

    os.chdir(zippath)
    files = os.listdir(aname)

    for i in files:
        file = os.path.join(aname, i)
        zipfd.write(file)
    zipfd.close()

    os.chdir(sfold)