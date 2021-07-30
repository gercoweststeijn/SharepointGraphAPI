############################################################# 
# auteur         versie      toelichting
# Weststeijn     1.0         creatie
#
# kopieer bestanden uit subfolder(s) van een bron directory
# naar een enkele doel directory 
#############################################################
import os, shutil, datetime, sys


#bron_map = 'C:/Temp/knipbestanden/'
bron_map = 'C:/Temp/knipbestanden/'

now = datetime.datetime.now()
ts  = now.strftime('%Y-%m-%d-%H_%M_%S')
result_file_name = 'c:/Temp/Bestanden_en_Path.csv'
result_file = open(result_file_name, "a", encoding='utf-8')

counter_all = 0
counter_cp = 0
for root, dirs, files in os.walk(bron_map): 
   for file in files:
      path_file = os.path.join(root,file) 
      result_file.write(file + '||'+path_file+'\n') 
      

result_file.close()
#print (str(counter_all) + ' bestanden tegen gekomen.')
#print (str(counter_cp) + ' bestanden gekopieerd controleer dat dit klopt met aantal in de doelmap')
