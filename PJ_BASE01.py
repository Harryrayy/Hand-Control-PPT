import win32com.client
from cvzone.HandTrackingModule import HandDetector
import cv2
import pyautogui
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox

# Global variable for selected PPT file
ppt_file = ""

# Function to open PowerPoint file
def open_ppt():
    global ppt_file
    ppt_file = filedialog.askopenfilename(filetypes=[("PowerPoint files", "*.pptx")])
    if ppt_file:
        file_label.config(text=f"Selected file: {ppt_file.split('/')[-1]}")
    else:
        file_label.config(text="No file selected")

# Function to start the presentation control program
def start_program():
    if not ppt_file:
        messagebox.showwarning("No file", "Please select a PowerPoint file before starting.")
        return

    Application = win32com.client.Dispatch("PowerPoint.Application")
    Presentation = Application.Presentations.Open(ppt_file)
    Presentation.SlideShowSettings.Run()

    # Parameters
    width, height = 900, 720
    gestureThreshold = 300

    # Camera Setup
    cap = cv2.VideoCapture(0)
    cap.set(3, width)
    cap.set(4, height)

    # Hand Detector
    detectorHand = HandDetector(detectionCon=0.8, maxHands=1)

    # Variables
    buttonPressed = False
    counter = 0
    imgNumber = 0
    delay = 25
    slideChangeFeedback = ""
    laserPointerActive = False

    # Get the screen size
    screenWidth, screenHeight = pyautogui.size()

    while True:
        # Get image frame
        success, img = cap.read()
        if not success:
            break

        
        img = cv2.flip(img, 1)

        
        hands, img = detectorHand.findHands(img)

        if hands and not buttonPressed:  # If hand detected
            hand = hands[0]
            cx, cy = hand["center"]
            lmList = hand["lmList"]
            fingers = detectorHand.fingersUp(hand)

            if cy <= gestureThreshold:
                if fingers == [1, 1, 1, 1, 1]:
                    slideChangeFeedback = "Next Slide"
                    buttonPressed = True
                    if imgNumber < Presentation.Slides.Count - 1:
                        Presentation.SlideShowWindow.View.Next()
                        imgNumber += 1  # Move to next slide
                elif fingers == [1, 0, 0, 0, 0]:  # Thumbs up
                    slideChangeFeedback = "Previous Slide"
                    buttonPressed = True
                    if imgNumber > 0:
                        Presentation.SlideShowWindow.View.Previous()
                        imgNumber -= 1  # Move to previous slide
                elif fingers == [0, 1, 0, 0, 0]:  # Index finger pointing
                    if not laserPointerActive:
                        pyautogui.hotkey('ctrl', 'l')  # Toggle laser pointer mode on
                        laserPointerActive = True
                    else:
                        pyautogui.hotkey('ctrl', 'l')  # Toggle laser pointer mode off
                        laserPointerActive = False

        if buttonPressed:
            counter += 1
            if counter > delay:
                counter = 0
                buttonPressed = False

        # Laser Pointer Logic
        if laserPointerActive and hands:
            indexFingerPos = lmList[8][:2]  # Index finger tip (landmark 8) x and y 
            # Convert from camera to screen
            mouseX = np.interp(indexFingerPos[0], [0, width], [0, screenWidth])
            mouseY = np.interp(indexFingerPos[1], [0, height], [0, screenHeight])
            pyautogui.moveTo(mouseX, mouseY)  # Move mouse to this position on the screen

            #the pointer position on the camera feed
            cv2.circle(img, (indexFingerPos[0], indexFingerPos[1]), 10, (0, 255, 0), -1)

        # Display slide number and feedback
        cv2.putText(img, f"Slide: {imgNumber + 1}/{Presentation.Slides.Count}", (10, 50),
                    cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 255, 255), 2)
        if slideChangeFeedback:
            cv2.putText(img, slideChangeFeedback, (300, 100), cv2.FONT_HERSHEY_SIMPLEX, 1.5, (0, 255, 0), 3)
            slideChangeFeedback = ""  # Reset after displaying for one frame

        # Show the image with UI
        cv2.imshow("Presentation Control", img)

        key = cv2.waitKey(1)
        if key == ord('q'):
            break

    # Release resources
    cap.release()
    cv2.destroyAllWindows()

# Function to exit the program
def exit_program():
    root.quit()

# Create the main window
root = tk.Tk()
root.title("PPT Intelligence")

# Create UI elements
file_label = tk.Label(root, text="No file selected", font=("Arial", 12))
file_label.pack(pady=100)

open_button = tk.Button(root, text="Open PowerPoint File", command=open_ppt, font=("Arial", 12), width=25)
open_button.pack(pady=5)

start_button = tk.Button(root, text="Start Presentation", command=start_program, font=("Arial", 12), width=25)
start_button.pack(pady=5)

exit_button = tk.Button(root, text="Exit", command=exit_program, font=("Arial", 12), width=25)
exit_button.pack(pady=5)

# Run the main loop
root.mainloop()
