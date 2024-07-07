import time
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.clock import Clock
from kivy.graphics.texture import Texture
from kivy.uix.image import Image
import cv2
from pyzbar.pyzbar import decode
import openpyxl
from openpyxl import Workbook

def find_working_camera_index():
    index = 0
    while True:
        cap = cv2.VideoCapture(index)
        if cap.isOpened():
            print(f"Camera found at index {index}")
            cap.release()
            return index
        cap.release()
        index += 1
        if index > 10:
            print("No available camera found within the first 10 indices.")
            return None

def write_to_excel(qr_data, additional_info):
    filename = 'attendance.xlsx'
    try:
        workbook = openpyxl.load_workbook(filename)
    except FileNotFoundError:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = 'QR Data'
        # Adjust the header to include 'QR Data' and 'Additional Info'
        sheet.append(['St_ID', 'Name', 'Date'])
    else:
        sheet = workbook['QR Data']
    # Write the qr_data and additional_info into separate columns
    sheet.append([qr_data, additional_info, time.strftime("%d/%m/%Y")])
    workbook.save(filename)
class CameraWidget(Image):
    def __init__(self, **kwargs):
        super(CameraWidget, self).__init__(**kwargs)
        self.capture = cv2.VideoCapture(find_working_camera_index() or 0)
        self.processed_qr_codes = set()
        Clock.schedule_interval(self.update, 1.0 / 30)

    def update(self, dt):
        ret, frame = self.capture.read()
        if ret:
            buf = cv2.flip(frame, 0).tobytes()
            texture = Texture.create(size=(frame.shape[1], frame.shape[0]), colorfmt='bgr')
            texture.blit_buffer(buf, colorfmt='bgr', bufferfmt='ubyte')
            self.texture = texture

            decoded_objects = decode(frame)
            for obj in decoded_objects:
                qr_data = obj.data.decode('utf-8')
                if qr_data not in self.processed_qr_codes:
                    print(f"QR Code Detected: {qr_data}")
                    self.processed_qr_codes.add(qr_data)
                    # Directly pass qr_data to write_to_excel, no splitting
                    write_to_excel(qr_data, "")
class AttendanceApp(App):
    def build(self):
        self.layout = BoxLayout(orientation='vertical')
        self.camera_widget = CameraWidget()
        self.submit_button = Button(text='Submit Attendance')
        self.info_label = Label(text='Scan QR Code for Attendance')

        self.submit_button.bind(on_press=self.submit_attendance)

        self.layout.add_widget(self.camera_widget)
        self.layout.add_widget(self.info_label)
        self.layout.add_widget(self.submit_button)

        return self.layout

    def submit_attendance(self, instance):
        print("Submit attendance logic goes here")

if __name__ == '__main__':
    AttendanceApp().run()