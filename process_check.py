from pywinauto import findwindows

# windows = Desktop(backend="uia").windows()
# print([w.window_text() for w in windows])


procs = findwindows.find_elements()

for proc in procs:
    print(f"{proc} / 프로세스 : {proc.process_id}")
    
# proc_list = procs['proc']