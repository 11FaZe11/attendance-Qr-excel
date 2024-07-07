# Attendance Tracking App

This application is designed to facilitate attendance tracking through QR code scanning. It utilizes the Kivy framework for the GUI, OpenCV for camera operations and QR code detection, and OpenPyXL for writing attendance data to an Excel file.

## Features

- **Camera Integration**: Automatically finds and uses the first working camera on the device to scan QR codes.
- **QR Code Scanning**: Scans QR codes in real-time and processes them without duplicates.
- **Excel Integration**: Records attendance data in an Excel file, creating a new file if one doesn't exist.

## Requirements

To run this application, you will need:

- Python 3.6 or higher
- Kivy
- OpenCV-Python
- Pyzbar
- OpenPyXL

## Installation

First, ensure that Python 3.6 or higher is installed on your system. Then, install the required packages using pip:

```bash
pip install kivy opencv-python pyzbar openpyxl
