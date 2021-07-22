
for i in range(1,50):
    f= open("c:/temp/test_"+str(i)+".txt","w+")
    for i in range(10):
        f.write("Dit is een test regel%d\r\n" % (i+1))
    f.close()   