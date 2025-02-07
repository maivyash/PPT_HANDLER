import win32com.client
import cv2
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import pyautogui
from PIL import Image, ImageTk
import mediapipe as mp

mp_hands = mp.solutions.hands
mp_draw = mp.solutions.drawing_utils

def select_ppt_file():
    file_path = filedialog.askopenfilename(title="Select PowerPoint File", filetypes=[("PowerPoint files", "*.pptx")])
    return file_path

def start_ppt():
    ppt_file = ppt_path.get()
    if ppt_file and os.path.exists(ppt_file):
        global Presentation
        Application = win32com.client.Dispatch("PowerPoint.Application")
        try:
            Presentation = Application.Presentations.Open(ppt_file)
            print(f"Opening PowerPoint file: {Presentation.Name}")
            Presentation.SlideShowSettings.Run()
            start_camera()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open PowerPoint: {e}")
    else:
        messagebox.showerror("Error", "No valid PowerPoint file selected or file does not exist!")

def get_finger_state(hand_landmarks):
    fingers = []
    finger_tips = [8, 12, 16, 20]
    for tip in finger_tips:
        fingers.append(hand_landmarks.landmark[tip].y < hand_landmarks.landmark[tip - 2].y)
    thumb_tip = hand_landmarks.landmark[4].x
    thumb_ip = hand_landmarks.landmark[3].x
    fingers.insert(0, thumb_tip < thumb_ip)
    return fingers

def start_camera():
    width, height = 640, 480
    cap = cv2.VideoCapture(0)
    cap.set(3, width)
    cap.set(4, height)
    hands = mp_hands.Hands(min_detection_confidence=0.9, min_tracking_confidence=0.9)
    buttonPressed = False
    counter = 0
    drawing = False
    delay_active = False
    slideShowView = Presentation.SlideShowWindow.View

    def reset_delay():
        nonlocal delay_active
        delay_active = False

    def next_slide():
        nonlocal delay_active
        if not delay_active:
            if slideShowView.Slide.SlideIndex == Presentation.Slides.Count:
                slideShowView.GotoSlide(1)  # Loop back to first slide
            else:
                slideShowView.Next()
            delay_active = True
            threading.Timer(1, reset_delay).start()

    def prev_slide():
        nonlocal delay_active
        if not delay_active:
            if slideShowView.Slide.SlideIndex == 1:
                slideShowView.GotoSlide(Presentation.Slides.Count)  # Loop back to last slide
            else:
                slideShowView.Previous()
            delay_active = True
            threading.Timer(1, reset_delay).start()

    def update_frame():
        nonlocal buttonPressed, counter, drawing
        success, img = cap.read()
        if not success:
            camera_label.after(10, update_frame)
            return
        img = cv2.flip(img, 1)
        img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        results = hands.process(img_rgb)
        
        if results.multi_hand_landmarks:
            for hand_landmarks in results.multi_hand_landmarks:
                mp_draw.draw_landmarks(img, hand_landmarks, mp_hands.HAND_CONNECTIONS)
                fingers = get_finger_state(hand_landmarks)
                cx, cy = int(hand_landmarks.landmark[8].x * width), int(hand_landmarks.landmark[8].y * height)
                
                if not buttonPressed:
                    if fingers == [0, 0, 0, 0, 1]:
                        print("🔙 Previous Slide")
                        buttonPressed = True
                        prev_slide()
                    elif fingers == [1, 0, 0, 0, 0]:
                        print("➡ Next Slide")
                        buttonPressed = True
                        next_slide()
                    elif fingers == [0, 1, 0, 0, 0]:
                        print("📍 Moving Pointer")
                        pyautogui.moveTo(cx * 2, cy * 2)
                        slideShowView.PointerType = 1  # Ensure red marker is active
                    elif fingers == [0, 1, 1, 0, 0]:
                        if not drawing:
                            print("✏️ Entering Marker Mode")
                            pyautogui.mouseDown()
                            slideShowView.PointerType = 1  # Ensure red marker is active
                            drawing = True
                        pyautogui.moveTo(cx * 2, cy * 2)
                    else:
                        if drawing:
                            print("❌ Stopping Drawing Mode")
                            pyautogui.mouseUp()
                            slideShowView.PointerType = 0
                            drawing = False

        if buttonPressed:
            counter += 1
            if counter > 5:
                counter = 0
                buttonPressed = False

        img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        img = Image.fromarray(img)
        img = ImageTk.PhotoImage(img)
        camera_label.img = img
        camera_label.config(image=img)
        camera_label.after(10, update_frame)
    
    update_frame()

root = tk.Tk()
root.title("PowerPoint Control System")
root.geometry("700x600")

label = tk.Label(root, text="Select PowerPoint file to start the presentation", font=("Arial", 14))
label.pack(pady=10)

ppt_path = tk.StringVar()

def choose_file():
    file = select_ppt_file()
    if file:
        ppt_path.set(file)
        file_button.config(state="disabled")
        start_button.config(state="normal")
        label.config(text=f"Selected File: {os.path.basename(file)}")

file_button = tk.Button(root, text="Choose PowerPoint File", font=("Arial", 12), command=choose_file)
file_button.pack(pady=10)

start_button = tk.Button(root, text="Start PPT & Control", font=("Arial", 12), state="disabled", command=start_ppt)
start_button.pack(pady=10)

camera_label = tk.Label(root)
camera_label.pack()

footer_label = tk.Label(root, text="Developed by Yash Gupta Contact: guptayash2005.yg@gmail.com", font=("Arial", 10))
footer_label.pack(side="bottom", pady=5)
root.mainloop()
