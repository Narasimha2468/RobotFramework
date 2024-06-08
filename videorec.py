import re
from datetime import datetime
import os
import subprocess
import cv2
import mss
import numpy as np
import asyncio

# class videorec:
#     def __init__(self,output_file) -> None:
#         self.output_file = output_file

async def start_video_recording(output_file, screen_size=(1920, 1080), fps=20.0):
    fourcc = cv2.VideoWriter_fourcc(*"mp4v")
    out = cv2.VideoWriter(output_file, fourcc, fps, screen_size)
    # stop_condition_file_path=r".\VideoRec_Info\TestSuiteInfo.json"
    stop_condition_txt = r".\VideoRec_Info\stop_recording.txt"
    with mss.mss() as sct:
        try:
            while not os.path.exists(stop_condition_txt):
                screenshot = sct.grab(sct.monitors[1])  # Capture the screen
                screenshot = np.array(
                    screenshot
                )  # Convert the screenshot to numpy array
                frame = cv2.cvtColor(
                    screenshot, cv2.COLOR_BGR2RGB
                )  # Convert the screenshot to a format compatible with OpenCV
                out.write(frame)  # Write the frame to the video file

                await asyncio.sleep(0.01)  # Allow other tasks to run asynchronously
        except asyncio.CancelledError:
            pass  # Stop recording when the task is cancelled
        finally:
            out.release()  # Release the VideoWriter and close the output file


# start_video_recording(r"C:\Users\Swaraj\Desktop\reports")
# start_recording_txt = r".\VideoRec_Info\StartRecording.txt"
# with open(start_recording_txt, "r") as rec:
#     savepath = rec.read()
#     # savepath = os.path.normpath(savepath)/
#     print((savepath))

savepath1 = r".\VideoRec_Info\testrecord.mp4"

# obj = videorec(savepath1)



# obj = videorec()
loop = asyncio.get_event_loop()
try:
    print("try")
    loop.run_until_complete(start_video_recording(savepath1))
    loop.close()
    print("except2")
except:
    # Handle manual interruption
    print("except1")
    pass
# finally:
#     loop.close()
