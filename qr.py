import cv2
from pyzbar import pyzbar

cap = cv2.VideoCapture(0)
qrCodeDetector = cv2.QRCodeDetector()

while True:

    ret, frame = cap.read()

    decodedText, points, _ = qrCodeDetector.detectAndDecode(frame)

    if points is not None:
 
        nrOfPoints = len(points)
        frame = cv2.polylines(frame, points.astype(int), True, (0, 255, 0), 3)

    if(decodedText != ""):
        decodedArray = decodedText.split("^")
        
        
    print(decodedText)    
        
     
 
    cv2.imshow("Image", frame)


    if(cv2.waitKey(1) == ord('q')):
        break

cap.release()
cv2.destroyAllWindows()