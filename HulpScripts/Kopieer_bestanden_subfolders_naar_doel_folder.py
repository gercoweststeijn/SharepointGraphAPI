############################################################# 
# auteur         versie      toelichting
# Weststeijn     1.0         creatie
#
# kopieer bestanden uit subfolder(s) van een bron directory
# naar een enkele doel directory 
#############################################################
import os, shutil, datetime, sys



bron_map = r'd:\box'
doel_map =r'd:\box_met_zips'

#now = datetime.datetime.now()
#ts  = now.strftime('%Y-%m-%d-%H_%M_%S')
#result_file_name = 'C:/temp/_ff_'+ts+'.csv'
#result_file = open(result_file_name, "a", encoding='utf-8')

counter_all = 0
counter_cp = 0
for root, dirs, files in os.walk(bron_map): 
   for file in files:
       
      counter_all = counter_all +1 
      doel_file = os.path.join(doel_map,file)
      path_file = os.path.join(root,file)
      print (doel_file)
      if not os.path.exists(doel_file):
         counter_cp = counter_cp +1
         
         shutil.copy2(path_file,doel_map) 
         print('GELUKT , ,' + file + ','+path_file+'\n') 
      else:
         print('FOUT , kunnen we niet schrijven want deze bestaat al ,'+file + ' , ' +path_file+'\n' )

#result_file.close()
print (str(counter_all) + ' bestanden tegen gekomen.')
print (str(counter_cp) + ' bestanden gekopieerd controleer dat dit klopt met aantal in de doelmap')
