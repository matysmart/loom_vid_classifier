import os
import shutil
import openpyxl
from sklearn.model_selection import train_test_split
from loom_dl import download_loom_video_to_file, remove_query_from_url

# Paths for train/val datasets
base_path = "loom_dataset"
train_path = os.path.join(base_path, "train")
val_path = os.path.join(base_path, "val")

# Function to create directories if they don't exist
def create_directories(colors):
    for color in colors:
        os.makedirs(os.path.join(train_path, color), exist_ok=True)
        os.makedirs(os.path.join(val_path, color), exist_ok=True)

def main():
    # Load the Excel file
    file_path = 'linksTestResultsGroundTruth.xlsx'  # Replace with your Excel file path
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active  # Or wb['SheetName'] if you want to specify the sheet

    # Extract videos and their colors
    videos_by_color = {'white': [], 'yellow': [], 'red': []}

    for row in sheet.iter_rows(min_row=2):  # Assuming headers are in the first row
        video_url = row[0].value  # Assuming video URLs are in the first column
        video_url = remove_query_from_url(video_url)
        color = row[0].fill.start_color.rgb
        
        if color == '00000000':  # Transparent as white
            videos_by_color['white'].append(video_url)
        elif color == 'FFFFFF00':  # Yellow
            videos_by_color['yellow'].append(video_url)
        elif color == 'FFFF0000':  # Red
            videos_by_color['red'].append(video_url)

    # Ensure output directories exist
    create_directories(videos_by_color.keys())

    # Split videos by 80% train, 20% val
    for color, videos in videos_by_color.items():
        train_videos, val_videos = train_test_split(videos, test_size=0.2, random_state=42)
        
        # Download and save videos in train/val directories
        for idx, video_url in enumerate(train_videos):
            save_path = os.path.join(train_path, color, f"video{idx+1}.avi")
            download_loom_video_to_file(video_url, save_path)
        
        for idx, video_url in enumerate(val_videos):
            save_path = os.path.join(val_path, color, f"video{idx+1}.avi")
            download_loom_video_to_file(video_url, save_path)

    print("Dataset created with train and val splits.")

if __name__ == "__main__":
    main()