import re
import time
import threading


def function1():
    # time.sleep(3)
    print("Function 1 is running")
    # 在这里实现function1的功能
 
def function2():
    # time.sleep(1)
    print("Function 2 is running")
    # 在这里实现function2的功能
 
def main():
    # 创建线程
    thread1 = threading.Thread(target=function1)
    thread2 = threading.Thread(target=function2)
 
    # 启动线程
    thread1.start()
    thread2.start()
 
    # 等待线程完成
    # thread1.join()
    # thread2.join()
    print("****finish****")

if __name__ == "__main__":
    '''
    date_pattern = r'\d{4}/\d{1,2}/\d{1,2}'
    date_text = "2021/11/15"
    print(re.match(date_pattern, date_text) is None)
    date_array = date_text.split("/")
    print(date_array)
    '''
    main()
    
