import datetime


#start,end,carpeta,reDimensionar,clearExif,fechar,iniNum,finNum,semanas,quincena,meses,fechaIni,fechaNumero,ajNum


start=1
end=68
carpeta = 'D:\\Media\\_Comic\\_PROCESS\\'.replace('\\', '/')

trash=carpeta+"imagenes basura"

#Ajusta el tamño de cada una de las imágenes
reDimensionar=True

#Limpiar metadatos de las imágenes
clearExif=True

#Controla imágenes giradas en Windows
controlGiro=False

#Recorta blancos
recortaBlanco=True

#Elimina Flyers
delFlyers=False

#Añade el la fecha al nombre del fichero
fechar=False
#Numero en la colección por el nombre del fichero
iniNum=0 #uno por detras
finNum=3 #justo
#recurrencia para la fecha
semanas=0
quincena=1
meses=0
#Fecha de uno de los números
# el jueves fechaIni= datetime.datetime(1978, 9, 20) 
fechaIni= datetime.datetime(1970, 7, 1) 
fechaNumero=1


# en negativo y menos 1,  Para corregir moviendo el numero que indica el fichero
ajNum=(0) 


#Recupera el título de un excel
datosFile='D:\\Media\\_Comic\\_PROCESS\\Data\\datos_el Jueves.xlsx'
titulo=False

#Compresión JPG
comprime=60