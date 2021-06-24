######################################################## 
# auteur         versie      toelichting
# Weststeijn     1.0         creatie
#
# kopieer bestanden uit subfolder van een 
#
########################################################
import os
import shutil

bron_map = 'C:/Temp/test/dl'
doel_map ='C:/Temp/test/WL'

counter = 0
for root, dirs, files in os.walk(bron_map): 
   for file in files:
      counter = counter +1 
      path_file = os.path.join(root,file)
      shutil.copy2(path_file,doel_map) 

print (str(counter) + ' bestanden tegen gekomen. Controleer dat dit aantal ook in de doelmap staat!')
