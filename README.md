# PowerPoint Gesture Control System


## 🚀 Overview
The **PowerPoint Gesture Control System** is an innovative application that allows users to control their PowerPoint presentations using simple hand gestures. Utilizing OpenCV, MediaPipe, and PyAutoGUI, this project enables seamless slide navigation and annotation without touching the keyboard or mouse.

## ✨ Features
✔ **Gesture-Based Slide Navigation**: Move to the next or previous slide using hand gestures.  
✔ **Pointer Mode**: Use your finger as a red marker to highlight content.  
✔ **Drawing Mode**: Activate drawing mode by raising two fingers and annotate slides in real-time.  
✔ **Looping Slideshow**: When reaching the last slide, the presentation restarts from the beginning.  
✔ **User-Friendly GUI**: Built using Tkinter for an easy-to-use interface.  
✔ **File Selection Dialog**: Select a PowerPoint file from the GUI before starting the slideshow.  

## 🛠 Installation
### 🔹 Prerequisites
Ensure you have the following installed:
- Python 3.x
- OpenCV (`cv2`)
- MediaPipe
- PyAutoGUI
- Pillow
- pywin32

### 🔹 Install Dependencies
Run the following command to install the required dependencies:
```sh
pip install opencv-python mediapipe pyautogui pillow pywin32
```sh

# PPT Gesture Control

## Install Dependencies

Run the following command to install the required dependencies:

```sh
pip install opencv-python mediapipe pyautogui pillow pywin32
```

## 🎬 Usage

### Run the Script

```sh
python ppt_gesture_control.py
```

### Select a PowerPoint File

1. Click on the **"Choose PowerPoint File"** button and select a `.pptx` file.
2. Once selected, the **"Start PPT & Control"** button will be enabled.

### Start Presentation

1. Click the **"Start PPT & Control"** button to launch the PowerPoint presentation.
2. The camera will activate to detect hand gestures.

## 🎮 Gesture Controls

- ✋ **Raise Thumb (only)** → Next Slide ⏭️
- 🤙 **Raise Pinky Finger (only)** → Previous Slide ⏮️
- ☝️ **Raise Index Finger (only)** → Move Red Marker 🎯
- ✌️ **Raise Index & Middle Fingers** → Enable Drawing Mode ✏️
- ✊ **Lower All Fingers** → Stop Drawing ❌

## 📸 Screenshots

![Gesture Control Demo](https://github.com/maivyash/PPT_HANDLER/blob/main/image.png)

## ⚠️ Known Issues

- 📷 **Camera Initialization Delay**: The camera might take a few seconds to initialize.
- 💡 **Lighting Conditions**: Ensure proper lighting for accurate gesture detection.

## 👨‍💻 Developer Information

- **Developed by**: Yash Gupta
- 📧 **Contact**: [guptayash2005.yg@gmail.com](mailto:guptayash2005.yg@gmail.com)

## 📜 License

This project is open-source under the **MIT License**.
