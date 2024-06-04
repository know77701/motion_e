class DashboardDto:
    def __init__(self, motion_window, motion_app, search_name, phone_number, start_sub_process_event, sub_process_done_event, btn_title):
        self.motion_window = motion_window
        self.motion_app = motion_app
        self.search_name = search_name
        self.phone_number = phone_number
        self.start_sub_process_event = start_sub_process_event
        self.sub_process_done_event = sub_process_done_event
        self.btn_title = btn_title
