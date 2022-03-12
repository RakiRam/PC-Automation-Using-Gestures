import math
import os
import time
import webbrowser
import wmi
from ctypes import cast, POINTER
import cv2
import numpy as np
from comtypes import CLSCTX_ALL
from cvzone.HandTrackingModule import HandDetector
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# volume metrics
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
volRange = volume.GetVolumeRange()
minVol = volRange[0]
maxVol = volRange[1]
vol = 0
volBar = 400
volPer = 0

# brightness metrics
minBrightness = 0
maxBrightness = 100

class HandDetectorClass():
    def __init__(self):
        super(HandDetectorClass, self).__init__()
        self.cap = cv2.VideoCapture(0)
        self.detector = HandDetector(detectionCon=0.8, maxHands=2)
        self.x = 0

    def readVideoFromCamera(self):
        while True:
            self.success, self.img = self.cap.read()
            self.hands, self.img = self.detector.findHands(self.img)
            if self.hands:
                self.InitializeHand()
                try:
                    self.DoActivity(self.x)
                except Exception as e:
                    print(e)
            cv2.imshow("Image", self.img)
            cv2.waitKey(1)

    def InitializeHand(self):
        self.hand1 = self.hands[0]
        self.lmList1 = self.hand1["lmList"]
        self.bbox1 = self.hand1["bbox"]
        self.centerPoint1 = self.hand1["center"]
        self.handType1 = self.hand1["type"]
        self.fingers1 = self.detector.fingersUp(self.hand1)

        if len(self.hands) == 2:
            self.hand2 = self.hands[1]
            self.lmList2 = self.hand2["lmList"]
            self.bbox2 = self.hand2["bbox"]
            self.centerPoint2 = self.hand2["center"]
            self.handType2 = self.hand2["type"]
            self.fingers2 = self.detector.fingersUp(self.hand2)
            try:
                self.length, self.info, self.img = self.detector.findDistance(self.centerPoint1, self.centerPoint2,
                                                                              self.img)
                self.x = int(self.length)
            except Exception as e:
                print("-----------")

        elif len(self.hands) == 1 and self.handType1 == "Left":
            self.VolumeControls(self.lmList1)

        elif len(self.hands) == 1 and self.handType1 == "Right":
            self.BrightnessControls(self.lmList1)

    def VolumeControls(self, lmList1):
        x1, y1 = self.lmList1[4][1], self.lmList1[4][2]
        x2, y2 = self.lmList1[8][1], self.lmList1[8][2]
        cx, cy = (x1 + x2) // 2, (y1 + y2) // 2

        cv2.circle(self.img, (x1, y1), 15, (255, 0, 255), cv2.FILLED)
        cv2.circle(self.img, (x2, y2), 15, (255, 0, 255), cv2.FILLED)
        cv2.line(self.img, (x1, y1), (x2, y2), (255, 0, 255), 3)
        cv2.circle(self.img, (cx, cy), 15, (255, 0, 255), cv2.FILLED)

        length = math.hypot(x2 - x1, y2 - y1)
        # print(length)
        # Hand range 50 - 300
        # Volume Range -65 - 0
        vol = np.interp(length, [50, 300], [minVol, maxVol])
        volBar = np.interp(length, [50, 300], [400, 150])
        volPer = np.interp(length, [50, 300], [0, 100])
        # print(int(length), vol)
        volume.SetMasterVolumeLevel(vol, None)

        if length < 50:
            cv2.circle(self.img, (cx, cy), 15, (0, 255, 0), cv2.FILLED)

        cv2.rectangle(self.img, (50, 150), (85, 400), (255, 0, 0), 3)
        cv2.rectangle(self.img, (50, int(volBar)), (85, 400), (255, 0, 0), cv2.FILLED)
        cv2.putText(self.img, f'{int(volPer)} %', (40, 450), cv2.FONT_ITALIC, 1, (255, 0, 0), 3)

    def BrightnessControls(self, lmList1):
        x11, y11 = self.lmList1[4][1], self.lmList1[4][2]
        x22, y22 = self.lmList1[8][1], self.lmList1[8][2]
        cx1, cy1 = (x11 + x22) // 2, (y11 + y22) // 2
        cv2.circle(self.img, (x11, y11), 15, (255, 0, 255), cv2.FILLED)
        cv2.circle(self.img, (x22, y22), 15, (255, 0, 255), cv2.FILLED)
        cv2.line(self.img, (x11, y11), (x22, y22), (255, 0, 255), 3)
        cv2.circle(self.img, (cx1, cy1), 15, (255, 0, 255), cv2.FILLED)
        length = math.hypot(x22 - x11, y22 - y11)
        brightness = np.interp(length, [50, 300], [minBrightness, maxBrightness])
        c = wmi.WMI(namespace='wmi')
        methods = c.WmiMonitorBrightnessMethods()[0]
        methods.WmiSetBrightness(brightness, 0)
        # print(int(brightness))
        BrightnessBar = np.interp(length, [50, 300], [400, 150])
        cv2.rectangle(self.img, (50, 150), (85, 400), (255, 0, 0), 3)
        cv2.rectangle(self.img, (50, int(BrightnessBar)), (85, 400), (255, 0, 0), cv2.FILLED)
        cv2.putText(self.img, f'{int(brightness)} %', (40, 450), cv2.FONT_ITALIC, 1, (255, 0, 0), 3)

    def DoActivity(self, x):
        if self.x in range(400, 500, 2):
            self.openWhatsApp()

    def openWhatsApp(self):
        webbrowser.open("https://web.whatsapp.com")


obj = HandDetectorClass()
obj.readVideoFromCamera()
