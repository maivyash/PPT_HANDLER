# PowerPoint Gesture Control System

## Overview
The **PowerPoint Gesture Control System** is an innovative application that allows users to control their PowerPoint presentations using simple hand gestures. Utilizing OpenCV, MediaPipe, and PyAutoGUI, this project enables seamless slide navigation and annotation without touching the keyboard or mouse.

## Features
- **Gesture-Based Slide Navigation**: Move to the next or previous slide using hand gestures.
- **Pointer Mode**: Use your finger as a red marker to highlight content.
- **Drawing Mode**: Activate drawing mode by raising two fingers and annotate slides in real-time.
- **Looping Slideshow**: When reaching the last slide, the presentation restarts from the beginning.
- **User-Friendly GUI**: Built using Tkinter for an easy-to-use interface.
- **File Selection Dialog**: Select a PowerPoint file from the GUI before starting the slideshow.

## Installation
### Prerequisites
Ensure you have the following installed:
- Python 3.x
- OpenCV (`cv2`)
- MediaPipe
- PyAutoGUI
- Pillow
- pywin32

### Install Dependencies
Run the following command to install the required dependencies:
```sh
pip install opencv-python mediapipe pyautogui pillow pywin32
