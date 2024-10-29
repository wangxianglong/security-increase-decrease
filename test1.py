from multiprocessing import Process
import time
 
def function1():
    # time.sleep(3)
    print("Function 1 is running")
    # 在这里实现function1的功能
 
def function2():
    # time.sleep(1)
    print("Function 2 is running")
    # 在这里实现function2的功能
 
def main():
    # 创建进程
    process1 = Process(target=function1)
    process2 = Process(target=function2)
 
    # 启动进程
    process1.start()
    process2.start()
 
    # 等待进程完成
    # process1.join()
    # process2.join()
    print("****finish****")
 
if __name__ == "__main__":
    main()