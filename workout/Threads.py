import threading
import v_1_0
# not needed ,available in _updating
def giveMeThreads(threadNo,totalPages):
    interval = int(totalPages/threadNo)
    list = []
    start = 1
    end = interval
    for i in range(1,threadNo):
        list.add( threading.Thread(target=_updating.iterate_pages,args=(start,end,)))
        start = end+1
        end+=interval
    # for last thread
    list.add(threading.Thread(target=_updating.iterate_pages, args=(start,totalPages ,)))
    return list

def startThreads(list):
    for t in list:
        t.start()

def joinThreads(list):
    for t in list:
        t.join()