import os

CacheFileName = "cache.txt"

def saveCache(personName, path):
    cache_d = os.path.join(path, CacheFileName)
    with open(cache_d, "a+") as f:
        # 读取所有行
        nameList = f.readlines()
        # 名字不存在则写入
        if personName not in nameList:
            f.write(personName)
            f.write("\n")
            f.flush()
        f.close()

def checkInCache(personName, path):
    cache_d = os.path.join(path, CacheFileName)
    if not os.path.exists(cache_d):
        f = open(cache_d, "w+")
        f.close()

    with open(cache_d, "r") as f:
        # 读取所有行
        nameList = f.readlines()
        f.close()
        if personName+"\n" not in nameList:
            return False
        else:
            return True
