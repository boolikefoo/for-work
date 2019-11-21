import time
start_time = time.time()



a = 100**100000
c = str(a)


print("--- %s seconds ---" % (time.time() - start_time), c.count('0') )