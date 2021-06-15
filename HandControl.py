import cv2
import time
import numpy as np
import HandTracking as ht
import math
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import win32api
from win32con import VK_MEDIA_PLAY_PAUSE, VK_MEDIA_NEXT_TRACK ,VK_MEDIA_PREV_TRACK

####################################
wCam, hCam = 1280, 720
####################################

cap = cv2.VideoCapture(0)
cap.set(3,wCam)
cap.set(4,hCam)
# cap.set(10,150)
pTime = 0

detector = ht.handDetector(detectionCon=0.8,maxHands=1)


devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
volRange = volume.GetVolumeRange()
minVol = volRange[0]
maxVol = volRange[1]
vol=0
volBar = 400
volPerc = 0
area = 0
colorVol = (255,0,0)


while True:
    success, img = cap.read()

    ### FIND HAND ######
    img = detector.findHands(img)
    lmList , bbox = detector.findPosition(img, draw=True)
    if len(lmList) != 0:

        ####### Filter Based on Size ##########
        area = (bbox[2] - bbox[0]) * (bbox[3] - bbox[1])//100
        if 250<area<1800:

            ####### Find distance between Index and Thumb ############

            length, img , lineInfo = detector.findDistance(4,8,img)

            ####### Convert Volume ############

            volBar = np.interp(length, [50, 200], [400, 150])
            volPerc = np.interp(length, [50, 200], [0, 100])

            ####### Reduce Resolution to make it smoother ############

            smoothness = 5
            volPerc = smoothness * round(volPerc/smoothness)

            ####### Check Fingers Up ############
            fingers = detector.fingersUp()

            ####### If Middle finger,Ring Finger and Pinky is Down ###########

            if not fingers[2] and not fingers[3] and not fingers[4] and fingers[0] and fingers[1] :
                volume.SetMasterVolumeLevelScalar(volPerc/100,None)
                cv2.circle(img, (lineInfo[4], lineInfo[5]), 15, (0, 255, 0), cv2.FILLED)
                colorVol = (0, 255, 0)
            else:
                colorVol = (255,0,0)

            #PLAY/PAUSE
            print(fingers)
            if fingers[0] and fingers[1] and fingers[4] and not fingers[2] and not fingers[3] :
                win32api.keybd_event(VK_MEDIA_PLAY_PAUSE, 0)
                time.sleep(0.3)
            ##NEXT TRACK
            elif fingers[1] and fingers[2] and not fingers[0] and not fingers[3] and not fingers[4]:
                win32api.keybd_event(VK_MEDIA_NEXT_TRACK, 0)
                time.sleep(0.3)

            ##PREVIOUS TRACK
            elif fingers[2] and fingers[3] and fingers[4] and not fingers[0]  and not fingers[1]:
                win32api.keybd_event(VK_MEDIA_PREV_TRACK, 0)
                time.sleep(0.3)


    ####### Draw ############

    cv2.rectangle(img,(50,150),(85,400),(255, 0, 0),3)
    cv2.rectangle(img, (50, (int(volBar))), (85, 400), (255, 0, 0), cv2.FILLED)
    cv2.putText(img, f'{int(volPerc)} %', (40, 450), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 0, 0), 3)

    cVol = int(volume.GetMasterVolumeLevelScalar()*100)
    cv2.putText(img, f'Current Volume : {int(cVol)} %', (800, 50), cv2.FONT_HERSHEY_COMPLEX, 1, colorVol, 3)

    ####### Frame Rate #############
    cTime = time.time()
    fps = 1 / (cTime-pTime)
    pTime = cTime

    cv2.putText(img,f'FPS:{int(fps)}',(40,50),cv2.FONT_HERSHEY_COMPLEX,1,(255,0,0),3)

    cv2.imshow("Img", img)
    if cv2.waitKey(10) & 0xFF == ord('q'):
        break
cv2.destroyAllWindows()
