import threading

def func(name,age):
    print(name,' is of age ',age)

#main code here
t1 = threading.Thread(target=func, args=('bala',17,))
t1.start()
t1.join()
t2 = threading.Thread(target=func,args=('hello',18,))
t2.start()

t2.join()
print("done")