import mmap
import contextlib
from datetime import datetime
import random
import time

# with contextlib.closing(mmap.mmap(-1, 1024, tagname='global_share_memory', access=mmap.ACCESS_WRITE)) as m:
with contextlib.closing(mmap.mmap(-1, 10, tagname='xran', access=mmap.ACCESS_WRITE)) as m:
    while True:
        m.seek(0)
        print(m.read(10))
    # for i in range(1, 100):
    #     m.seek(0)
    #     nt = datetime.now().strftime('%Y-%m-%d %H:%M:%S %f')
    #     m.write((nt).encode())
    #     m.flush()
    #     # t = random.randrange(0,10,1)
    #     print(datetime.now(), "msg " + str(nt))
    #
        time.sleep(2)