import logging
import math
import os
import shutil
import stat
import subprocess
import sys
from contextlib import closing
from pathlib import Path
from zipfile import ZipFile

import exifread

#pip install PyMuPDF Pillow
import fitz
import numpy as np
#pip install patool
import patoolib
import rarfile
from dateutil.relativedelta import relativedelta

from PIL import Image
from PIL import ImageFile
ImageFile.LOAD_TRUNCATED_IMAGES = True

import filetype
# pip install Send2Trash
from send2trash import send2trash

import Vars
import Data_manage

import TrimImageWhite

from pathvalidate import sanitize_filepath

#Log*********************************************************

logging.basicConfig(filename=Vars.carpeta+'python.log', encoding='utf-8', level=logging.ERROR)
logging.basicConfig(format='%(asctime)s %(levelname)s:%(message)s', level=logging.ERROR)

#Log*********************************************************

imagenSal=(".jpg",".png")
comic_recuperado=""

"""fechas*******************************************************
mañana      = Vars.fechaIni + relativedelta(days=+1)
semana      = Vars.fechaIni + relativedelta(weeks=-184)
next_month  = Vars.fechaIni + relativedelta(months=+1)
#fechas*******************************************************
"""


def calculaFecha(numero):
    global fechaIni,semanas,quincena,meses
    
    comicDate=Vars.fechaIni
    try:
        if Vars.semanas > 0:
            comicDate=Vars.fechaIni+relativedelta (weeks=+ (numero-Vars.fechaNumero)*Vars.semanas)
        elif Vars.quincena > 0:
                comicDate=Vars.fechaIni+relativedelta (months=+ math.trunc((numero-Vars.fechaNumero)/2)*Vars.quincena)
                if (numero % 2) == 0:
                    comicDate=comicDate+relativedelta (days=+15)  
        elif Vars.meses>0:
                comicDate=Vars.fechaIni+relativedelta (months=+ (math.trunc(numero-Vars.fechaNumero)*Vars.meses))                            
        return comicDate
    except Exception as inst:
        print(type(inst))
        print(inst.args)
        print(inst) 
        

def renombraFichero(count,comic,tipo):
    global iniNum,finNum,ajNum,carpeta
    try:
        fichero=os.path.basename(comic)

        if Vars.titulo or Vars.fechar:
            numero=int(fichero[Vars.iniNum:Vars.finNum])+Vars.ajNum #reajuste de numero de la colección
        
        tit=""
        if Vars.titulo:
            tit=Data_manage.df["Titulo"].values[numero-1]
            #tit= str(tit[3:-1].lower()) # Ajuste del texto
            tit= str(tit.lower()) # Ajuste del texto
            tit=tit.strip()
            tit= "-"+ tit.capitalize()      

        comicDate=""
        if Vars.fechar:            
            comicDate=calculaFecha(numero)   
            comicDate=comicDate.strftime("%d-%m-%Y")             
        
        path=Vars.carpeta+comicDate+"~"+str(count)+"#"+fichero[:-4] +str(tit)+"."+tipo
        newPath=sanitize_filepath(path)
        os.rename(Vars.carpeta+fichero, newPath)
        

    except Exception as inst:
        print(type(inst))
        print(inst.args)
        print(inst)    


def trataCbr(comic):
    global comic_recuperado
    logging.info( 'F >  '+comic)

    try:
        with closing(rarfile.RarFile (comic, 'r')) as archivo:
            borrados=0
            for borrar in archivo.infolist():
                if borrar.filename[-1:] =='/':
                    borrados=+1
                    continue
                if borrar.filename[-1:] !='/' and borrar.filename[-4:].lower() !='.jpg' and borrar.filename[-5:].lower() !='.jpeg'and borrar.filename[-4:].lower() !='.png':
                    logging.info( 'd >  '+borrar.filename)
                    borrados=+1
                    subprocess.call(["rar", "d", borrar.volume_file, borrar.filename.replace('/', '\\')], shell=True)
            count = len(archivo.namelist())-borrados
            
        logging.info( 'P >  '+str(count))
        renombraFichero(count, comic,'cbr')

    except Exception as inst:
        logging.error( '**E** > Zip >  '+comic)
        print(type(inst))
        print(inst.args)
        print(inst)        

        if (comic_recuperado!=comic):            
            comic_recuperado=comic
            trataCbz(comic)
        else:    
            logging.error( '**E******** > fichero no tratable >  '+comic)
            print(type(inst))
            print(inst.args)
            print(inst)

def trataCbz(comic):
    global comic_recuperado
    logging.info('F >  '+comic)

    reEnvio=False
    try:
        with closing(ZipFile(comic, mode='a')) as archivo:
            borrados=0
            if len (archivo.infolist())>0:                    
                for borrar in archivo.infolist():                                
                    if borrar.filename[-1:] =='/':
                        borrados=+1
                        continue
                    if borrar.filename[-1:].lower() !='/' and borrar.filename[-4:].lower() !='.jpg' and borrar.filename[-5:].lower() !='.jpeg'and borrar.filename[-4:].lower() !='.png':
                        logging.info( 'd >  '+borrar.filename)
                        borrados=+1
                        archivo.remove(borrar.filename)                                
                count = len(archivo.namelist())-borrados
                logging.info( 'P >  '+str(count)+'\n')
            else:
                reEnvio=True
        
        if reEnvio == False:
            renombraFichero(count, comic, 'cbz')
                            
        if (reEnvio and comic_recuperado!=comic):            
            comic_recuperado=comic
            trataCbr(comic)
        else:    
            logging.error( '**E******** > fichero no tratable >  '+comic)
            print(type(inst))
            print(inst.args)
            print(inst)
                
    except Exception as inst:
        logging.error( '**E** > Rar >  '+comic)
        print(type(inst))
        print(inst.args)
        print(inst)
        
        if (comic_recuperado!=comic):
            comic_recuperado=comic
            trataCbr(comic)
        else:    
            logging.error( '**E******** > fichero no tratable >  '+comic)
            print(type(inst))
            print(inst.args)
            print(inst)

def trataPdf(comic):
    logging.info( 'PDF >  '+str(comic)+'\n')
    imgdir = comic[:-4] 
    

    if not os.path.exists(imgdir):
        os.mkdir(imgdir)

    doc = fitz.open(comic)
    for page in doc:
        pix = page.get_pixmap(matrix=fitz.Identity, dpi=150, 
                            colorspace=fitz.csRGB, clip=None, annots=False)
        
        pagImagen=imgdir+"\\pag_%s.%s" %(str(page.number).zfill(4), imagenSal[0])

        pix.save(pagImagen, quality=Vars.comprime)  # save file
        if Vars.reDimensionar:
            SanitizeImage(pagImagen)
    doc.close()   

    creaCBR(imgdir)

   
# Óptimo 
# JPEG, progressive, quality: 82, subsampling ON (2x2)
# 1188 x 1656  Pixels (1.97 MPixels) (0.71)
# 20.1 x 28.0 cm; 7.92 x 11.04 inches
# 16,7 Million   (24 BitsPerPixel)
def SanitizeImage(pagImagen):
    
    #Eliminar flyers
    if Vars.delFlyers:
        if Data_manage.crcCalc(pagImagen) in Data_manage.basura:      
            Data_manage.moveFilePlus(pagImagen, Vars.trash)
            return
        
    imagePath=Path(pagImagen)
    
    #Elimina flag solo lectura
    os.chmod(imagePath, stat.S_IWRITE)

    #Recorta márgenes de las páginmas que están en blanco 
    if Vars.recortaBlanco:
        #Parche para solucionar bug en CV2 con caracteres extendidos
        shutil.move(imagePath, "a.jpg")
        TrimImageWhite.trimWhite("a.jpg", "b.jpg")        
        shutil.move("b.jpg", imagePath)
        os.remove("a.jpg")

    f = open(pagImagen, 'rb')

    #Soluciona problemas de giros dentro de windows
    if Vars.controlGiro:
        tags = exifread.process_file(f)
        for tag in tags.keys():
            #print ("Key: %s, value %s" % (tag, tags[tag])) 
            if tag not in ('JPEGThumbnail', 'TIFFThumbnail', 'Filename', 'EXIF MakerNote') and str(tags[tag]).startswith("Rotated "):         
                image = Image.open(pagImagen)
                data = list(image.getdata())
                image_without_exif = Image.new(image.mode, image.size)
                image_without_exif.putdata(data)

                match str(tags[tag]):
                    case "Rotated 180":
                        image.rotate(180).save(pagImagen, quality=Vars.comprime)
                    case "Rotated 90 CCW":
                        image.rotate(90, expand=True).save(pagImagen, quality=Vars.comprime)   
                    case "Rotated 90 CW": 
                        image.rotate(-90, expand=True).save(pagImagen, quality=Vars.comprime)
                print ("rotado ")
                logging.info( 'Rotado >  '+pagImagen)              
                image.close()
                break
    

    image = Image.open(pagImagen)
    try:
        ancho=math.trunc((image.size[0]/image.info['dpi'][0])*25.4)
        alto=math.trunc((image.size[1]/image.info['dpi'][1])*25.4)

        print(str(ancho) +"x"+ str(alto))
        print(str(image.size)+", "+str(image.info['dpi']))
        reducir = np.where(image.info['dpi'][0] or image.info['dpi'][1] > 144, True, False)

        #cálculo a 144dpi
        wpi=math.trunc((ancho/25.4)*144)
        hpi=math.trunc((alto/25.4)*144)

    except Exception as exc:
        
        #La imagen no trae info sobre DPI
        ratio=image.size[0]/image.size[1]
        if ratio > 1: #apaisado para A4
            wpi=math.trunc((420/25.4)*144)
        else: #retrato para A4
            wpi=math.trunc((210/25.4)*144)
        hpi= math.trunc((wpi*image.size[1])/image.size[0])

        reducir = np.where(image.size[0]>hpi or image.size[0]> wpi, True, False)

        

    #Redimensiona la imagen
    if Vars.reDimensionar and reducir:

        image.thumbnail((wpi, hpi))
        image.save(pagImagen, quality=Vars.comprime)
       

    #Limpiar metadatos de las imágenes
    if Vars.clearExif:
        data = list(image.getdata())

        image_without_exif = Image.new(image.mode, image.size)
        image_without_exif.putdata(data)
    
    image.close()


    

def creaCBR(imgdir):
    
    for file_n in os.listdir(imgdir):
         if file_n.endswith(imagenSal):
            os.system('rar m -ep -m5 "%s.cbr" "%s"' %(imgdir, imgdir+"\\"+file_n)) 

    os.rmdir(imgdir)   
        

def reProcesa(comic):
      
    imgdir = comic[:-4]

    if not os.path.exists(imgdir):
            os.mkdir(imgdir)
    
    try:
        patoolib.extract_archive(comic, outdir=imgdir)  
    except Exception as inst:
        logging.error(type(inst))
        logging.error(inst.args)
        logging.error(inst) 
        sys.exit(1) 

    send2trash( Path(comic) )   
    print("Check images")

    for root, dirs, files in os.walk(imgdir):

        path = root.split(os.sep)
        print((len(path) - 1) * '---', os.path.basename(root))
        
        if  os.path.basename(root)=="__MACOSX":
            shutil.rmtree(root, ignore_errors=False, onerror=None)
            continue
            
        for imagen in files:
            print(len(path) * '---', imagen)
            imagen1=root+"/"+imagen
            
            if imagen1.lower().endswith(imagenSal):            
                if imagen1.lower().endswith(imagenSal[1]):
                    imagen2=imagen1[:-4] +str('.jpg')

                    im1 = Image.open(imagen1)
                    im1.save(imagen2, quality=Vars.comprime)
                    os.remove(imagen1)
                    
                    imagen1=imagen2
                    im1.close

                SanitizeImage(imagen1)

    print ("Check images")
    
    os.system('rar m -r -ep1 -m5 -df "%s.cbr" "%s"' %(imgdir, imgdir+"/*.*"))         
    shutil.rmtree(imgdir, ignore_errors=False, onerror=None)

def normFichero (fichero):
    comic=os.path.join(Vars.carpeta, fichero )
    pre, ext = os.path.splitext(comic)   
    pre=pre.rstrip()

    kind= filetype.guess(Vars.carpeta+fichero) 
    if fichero == "python.log" or fichero == "listado_faltas.txt":
        print('Fichero log')
        return None
    
    elif kind is None:
        print('Fichero NO identificado!')
        sys.exit(1) 
    else:
        print('File extension: %s' % kind.extension)
        print('File MIME type: %s' % kind.mime)

    if kind.mime=="application/x-rar-compressed":
        os.rename(comic, pre + ".cbr")
        ext="cbr"
    elif  kind.mime== "application/zip":
        os.rename(comic, pre + ".cbz") 
        ext="cbz"
    elif  kind.mime== "application/x-7z-compressed":
        os.rename(comic, pre + ".cbz") 
        ext="cbz"        
    elif kind.mime== "application/pdf":
        os.rename(comic, pre + ".pdf")   
        ext="pdf"
    
    return pre+"."+ext
  

for fichero in os.listdir(Vars.carpeta):
    
    #fichero=fichero[:-4].rstrip()+fichero [len(fichero)-4:]
    if os.path.isfile(Vars.carpeta+fichero):
        comic= normFichero(fichero)
    else:
        continue 

    if comic == None or not os.path.isfile(comic):
        continue

    if   comic.endswith('.cbr'):        
        if Vars.reDimensionar:
            reProcesa(comic) 
        trataCbr(comic)       
    
    elif comic.endswith('.cbz'):
        if Vars.reDimensionar:
            reProcesa(comic) 
        trataCbr(comic[:-4]+".cbr")                        
    
    elif comic.endswith('.pdf'):         
        trataPdf(comic)
        send2trash( Path(comic) ) 
        
                 