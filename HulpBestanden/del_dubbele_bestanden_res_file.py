############################################################# 
# auteur         versie      toelichting
# Weststeijn     1.0         creatie
#
# delelete bestanden
#############################################################
import os, shutil, datetime, sys
import csv

now = datetime.datetime.now()
ts  = now.strftime('%Y-%m-%d-%H_%M_%S')
result_file_name = 'c:/Temp/ResultFile_del_'+ts+'.csv'
result_file = open(result_file_name, "a", encoding='utf-8')

bron_map ='C:/Temp/rwzibron/'
filename = r'c:\Temp\ResultFile_copy_2021-07-27-08_30_26.csv'

counter_all = 0
counter_del = 0

with open(filename, 'r', encoding='utf-8') as csvfile:
    datareader = csv.reader(csvfile)
    for row in datareader:
        counter_all = counter_all+1
        if row[0] == 'FOUT ':
            target_file = os.path.join(bron_map,row[2].strip())
            print (target_file)
            if os.path.exists(target_file):
                counter_del = counter_del +1
                os.remove(target_file) 
                result_file.write('GELUKT , ,' + target_file + '\n') 
            else:
                result_file.write('FOUT , kunnen we niet deleten, is er niet meer ,'+target_file +'\n' )


print ('behandeld: ' + str(counter_all))
print ('verwijderd: ' + str(counter_del))

result_file.write('behandeld: ' + str(counter_all))
result_file.write('verwijderd: ' + str(counter_del))
result_file.close()