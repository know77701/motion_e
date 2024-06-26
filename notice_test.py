from pywinauto import findwindows, Desktop, keyboard
from pywinauto import application
import time

app_win32 = application.Application(backend='win32')
app_uia = application.Application(backend='uia')

# app_win32.start("C:\\Motion\\Motion_E\\Motion_E.exe")
# print("실행 성공")
# time.sleep(10)

# cvt = app_win32.window(title='로그인')
# print("타이틀 확인 성공 ")

# cvt.child_window(title='로그인').click()
# print("로그인 버튼 선택 성공")
# # time.sleep(30)

app_uia.connect(path="C:\\Motion\\Motion_E\\Motion_E.exe")
print("커넥트 성공")
time.sleep(5)

uia_cvt = app_uia.window(title='모션.ver[1.0.0.237]ipc://Motion_E-64bit')

uia_cvt.child_window(auto_id='notice-content',control_type='Edit').type_keys('TEST{ENTER}')
print("공지등록 완료")
time.sleep(3)

uia_cvt.child_window(title='닫기',control_type='Button',found_index=0).click()
# uia_cvt.child_window(auto_id='notice-content',control_type='Edit').type_keys('{ENTER}')
time.sleep(2)
print("공지 삭제 클릭")


# uia_cvt.child_window(auto_id='radButton1',control_type='Button').click()
# time.sleep(2)

rad = app_uia.window(auto_id="RadMessageBox")             
radBtn = rad.child_window(auto_id="radButton1", control_type="Button")
radBtn.click()
print("공지삭제 성공")



# class notice :
    
    # 공지사항 쓰기
    # def notice_write(create) :
    #     try:
    #         app_uia.child_window(auto_id='notice-content',control_type='Edit').type_keys(create)
    #         time.sleep(1)
    #         keyboard.send_keys('{ENTER}')
    #         print("공지사항 생성 완료")
        
    #     except Exception as e :
    #         keyboard.send_keys('{F5}')
    #         print("공지생성실패",e)
        
        
    
    # 공지사항 삭제
    # def notice_delete() :
    


# class motion_start:
    
# # uia부터 체크하고, 있으면 connect / 없으면 win32 체크 후 진행
#     def check(proc) :
#         windows_uia = Desktop(backend='uia').windows()
        
#         for windows_uia in windows :
#             try :
#                 window_text = windows_uia.window_text()
#                 if proc in window_text :
                    
                    
                
    

        
    # @staticmethod
    # def app_start:




















