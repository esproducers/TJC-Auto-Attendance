import cv2
import numpy as np

print("Opening Raw Camera Feed Viewer...")
print("If the popup window is purely black, please check your physical laptop privacy slider or F8 key!")

cap = cv2.VideoCapture(0)
cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)

while True:
    ret, frame = cap.read()
    if ret and frame is not None:
        mean_brightness = np.mean(frame)
        
        # Add a text overlay to explain what we are seeing
        cv2.putText(frame, f"Raw Sensor Feed. Brightness: {mean_brightness:.1f}/255", 
                    (20, 50), cv2.FONT_HERSHEY_SIMPLEX, 1.0, (0, 0, 255), 2)
        cv2.putText(frame, "If this entire screen is black, your physical", 
                    (20, 100), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 0, 255), 2)
        cv2.putText(frame, "HP lens cover is closed, or the F8 camera kill switch is ON.", 
                    (20, 140), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 0, 255), 2)
        cv2.putText(frame, "Press 'Q' to close this window.", 
                    (20, 200), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 255, 0), 2)
                    
        cv2.imshow("DIAGNOSTIC CAMERA VIEWER", frame)
        
    if cv2.waitKey(30) & 0xFF == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()
