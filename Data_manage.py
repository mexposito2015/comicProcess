
#pip install xlrd
#pip3 install pandas
#pip install openpyxl

#  https://sparkbyexamples.com/pandas/pandas-get-cell-value-from-dataframe/


import os
import shutil
import zlib
import pandas as pd
import Vars


if Vars.titulo:
    df = pd.read_excel(Vars.datosFile)


    print(df)


    print("tests **************************")
    #print(df.loc[[8]])

    #print(df.iloc[8]['Titulo'])

    tit=df["Titulo"].values[15]

    tit= str(tit[4:-1].lower())
    print (tit.capitalize())


#Listado imágenes basura

def crcCalc(fichero):
    buffersize = 65536

    with open(fichero, 'rb') as afile:
        buffr = afile.read(buffersize)
        crcvalue = 0
        while len(buffr) > 0:
            crcvalue = zlib.crc32(buffr, crcvalue)
            buffr = afile.read(buffersize)

    print(format(crcvalue & 0xFFFFFFFF, '08x')) # a509ae4b

    return crcvalue

#Devuelve una listas de los CRC de imágenes que son basura en los comics
if Vars.delFlyers:
    basura=[]
    for fichero in os.listdir(Vars.trash):
        crc=crcCalc(Vars.trash+"\\"+fichero)
        if crc not in basura:
            basura.append(crc)
    


def moveFilePlus (file, movdir):
    
    name= os.path.splitext(os.path.basename(file))[0]
    extension= os.path.splitext(file)[1]
    fileOut= movdir+"/"+name+extension
    
    if not os.path.exists(os.path.join(fileOut)):
        shutil.move(file, fileOut)
    else:
        i = 1
        while True:
            new_name = movdir +"/" +name + "_" + str(i) + extension
            if not os.path.exists(new_name):
                shutil.move(file, new_name)                
                break 
            i += 1