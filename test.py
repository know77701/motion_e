import multiprocessing
import time

def main_task(quit_event, motion_window, event):
    # 여기에 main_task의 작업 내용을 추가
    print(f"Main task started with window: {motion_window}")
    
    # 예를 들어, 여기에 버튼 클릭 등의 작업을 수행
    print('테스트')    

    # 작업이 완료되면 이벤트를 설정
    time.sleep(1)  # 작업 시뮬레이션을 위한 지연 시간
    print("Main task completed. Setting the event.")
    event.set()
    print("ㅁㄴㅇㅁㄴㅇㅁㄴ")
    # quit_event 설정
    quit_event.wait()
    print("aaaa")

def sub_task(quit_event, event):
    # 이벤트가 설정될 때까지 대기
    event.wait()
    
    # 이벤트가 설정되면 서브 작업 수행
    time.sleep(2)
    print("Sub task started after main task completion.")
    
    # 여기에 sub_task의 작업 내용을 추가
    print("Sub task completed.")
    
    # quit_event 설정
    quit_event.set()

def run_tasks():
    quit_event = multiprocessing.Event()
    event = multiprocessing.Event()
    motion_window = "YourMotionWindowID"  # 여기에 실제 window ID를 사용

    process1 = multiprocessing.Process(target=main_task, args=(quit_event, motion_window, event))
    process2 = multiprocessing.Process(target=sub_task, args=(quit_event, event))

    process1.start()
    process2.start()

    process1.join()
    process2.join()

if __name__ == '__main__':
    run_tasks()