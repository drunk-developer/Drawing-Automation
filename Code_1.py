# CODE 1
#CONVERT ALL PPT TO PPTX



import os
import win32com.client
 
def convert_ppt_to_pptx(input_folder):
    ppt_files = [f for f in os.listdir(input_folder) if f.lower().endswith('.ppt')]
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = 1
 
    for ppt_file in ppt_files:
        full_path = os.path.join(input_folder, ppt_file)
        pptx_path = os.path.join(input_folder, ppt_file[:-4] + '.pptx')
        print(f"Converting {ppt_file} to {pptx_path}...")
        presentation = powerpoint.Presentations.Open(full_path, WithWindow=False)
        presentation.SaveAs(pptx_path, FileFormat=24)  # 24 = pptx format
        presentation.Close()
    powerpoint.Quit()
    print("All .ppt files converted to .pptx.")
 
# Usage: Specify your folder
convert_ppt_to_pptx(r"C:\Users\INPUT PPTS")