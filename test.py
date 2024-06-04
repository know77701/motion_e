import multiprocessing
import time


def main_process_func(start_sub_process_event, sub_process_done_event):
    print("Main process started")
    # 메인 프로세스의 작업을 시뮬레이션
    for i in range(2):
        print(f"Main process working... step {i+1}")
        time.sleep(1)  # 작업을 시뮬레이션하는 대기 시간

    # 첫 번째 서브 프로세스 호출
    print("Main process completed specific task, starting sub process")
    start_sub_process_event.set()  # 서브 프로세스를 시작하도록 이벤트 설정
    sub_process_done_event.wait()  # 서브 프로세스가 완료될 때까지 대기

    print("Sub process completed, Main process resumes")
    sub_process_done_event.clear()
    start_sub_process_event.clear()

    # 메인 프로세스의 추가 작업을 시뮬레이션
    for i in range(2):
        print(f"Main process working... step {i+1}")
        time.sleep(1)  # 작업을 시뮬레이션하는 대기 시간

    # 두 번째 서브 프로세스 호출
    print("Main process completed specific task again, starting sub process")
    start_sub_process_event.set()
    sub_process_done_event.wait()

    print("Sub process completed again, Main process resumes")
    sub_process_done_event.clear()
    start_sub_process_event.clear()

    print("Main process completed")


def sub_process_func(start_sub_process_event, sub_process_done_event, run_number):
    while True:
        start_sub_process_event.wait()  # 이벤트가 설정될 때까지 대기
        print("Sub process started")
        # 서브 프로세스의 작업을 시뮬레이션
        for i in range(run_number):
            print(f"Sub process working... step {i+1}")
            time.sleep(1)  # 작업을 시뮬레이션하는 대기 시간
        print("Sub process completed")
        sub_process_done_event.set()  # 서브 프로세스 완료 이벤트 설정
        start_sub_process_event.clear()  # 다음 작업을 위해 이벤트 재설정


if __name__ == '__main__':
    start_sub_process_event = multiprocessing.Event()
    sub_process_done_event = multiprocessing.Event()

    main_process = multiprocessing.Process(
        target=main_process_func, args=(start_sub_process_event, sub_process_done_event))
    sub_process = multiprocessing.Process(
        target=sub_process_func, args=(start_sub_process_event, sub_process_done_event, 1))

    main_process.start()
    sub_process.start()

    main_process.join()
    sub_process.terminate()  # 서브 프로세스를 종료합니다.

    print("Both processes have completed")