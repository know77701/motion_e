def factorial(n):
    # 기본 조건: n이 0이면 1을 반환
    if n == 0:
        return 1
    else:
        # 재귀 호출: n * (n-1)의 팩토리얼
        result = n * factorial(n - 1)
        print(f"factorial({n}) = {result}")
        return result


print(factorial(5))  # 출력: 120
