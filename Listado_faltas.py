import logging
import os

import Vars

#Log*********************************************************

logging.basicConfig(filename=Vars.carpeta+'listado_faltas.txt', encoding='utf-8', level=logging.DEBUG)
logging.basicConfig(format='%(asctime)s %(levelname)s:%(message)s', level=logging.DEBUG)

#Log*********************************************************

numeros= []
faltan=[]
for fichero in os.listdir(Vars.carpeta):
    if fichero=='listado_faltas.txt':
        continue
    numeros.append(int(fichero[Vars.iniNum:Vars.finNum]))

for numero in range(Vars.start, Vars.end+1):
    try:
        numeros.remove(numero)
    except ValueError:
        faltan.append(numero)
        pass  # No está    

faltan.sort()

for n in faltan:
    logging.info( "Falta "+ str(n))

logging.info( Vars.carpeta+"\n")
logging.info( "Faltan "+str(len(faltan))+" comics")