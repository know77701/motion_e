import multiprocessing


def worker(num):
    print('Worker', num)


if __name__ == '__main__':
    # 생성한 프로세스들을 저장할 리스트
    processes = []

    # 2개의 프로세스 생성
    for i in range(2):
        p = multiprocessing.Process(target=worker, args=(i,))
        processes.append(p)
        p.start()

    # 메인 프로세스에서 다른 작업 수행
    print('Main process is doing something else...')

    # 생성한 프로세스들이 종료될 때까지 기다림
    for process in processes:
        process.join()

    print('All processes have finished.')
