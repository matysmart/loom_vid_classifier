import argparse
import json
import urllib.request
from moviepy.video.io.VideoFileClip import VideoFileClip
import os
import pandas as pd
import shutil
import openpyxl
from sklearn.model_selection import train_test_split

# Paths for train/val datasets
base_path = "loom_dataset"
train_path = os.path.join(base_path, "train")
val_path = os.path.join(base_path, "val")

def fetch_loom_download_url(id):
    request = urllib.request.Request(
        url=f"https://www.loom.com/api/campaigns/sessions/{id}/transcoded-url",
        headers={},
        method="POST",
    )
    response = urllib.request.urlopen(request)
    body = response.read()
    content = json.loads(body.decode("utf-8"))
    url = content["url"]
    return url


def download_loom_video(url, filename):
    urllib.request.urlretrieve(url, filename)


def trim_video(input_file, output_file, start_time=1):
    with VideoFileClip(input_file) as video:
        trimmed_video = video.subclip(0, start_time)  # Trim the first 1 second
        trimmed_video.write_videofile(output_file, codec="mpeg4")


def parse_arguments():
    parser = argparse.ArgumentParser(
        prog="loom-dl", description="script to download loom.com videos"
    )
    parser.add_argument(
        "url", help="Url of the video in the format https://www.loom.com/share/[ID]"
    )
    parser.add_argument("-o", "--out", help="Path to output the file to")
    arguments = parser.parse_args()
    return arguments


def extract_id(url):
    return url.split("/")[-1]

def delete_file_if_exists(file_path):
    if os.path.isfile(file_path):
        os.remove(file_path)
        print(f"{file_path} has been deleted.")
    else:
        print(f"{file_path} does not exist.")

def download_loom_video_to_file(url_str, trimmed_filename):
    # url_str = "https://www.loom.com/share/fb8f4b80277249a681087be933bf546c"
    id = extract_id(url_str)

    url = fetch_loom_download_url(id)
    filename = f"{id}.mp4"
    print(f"Downloading video {id} and saving to {filename}")
    download_loom_video(url, filename)

    # Trim the first second of the downloaded video
    print(f"Trimming video {filename} to {trimmed_filename}")
    trim_video(filename, trimmed_filename)

    # Example usage
    delete_file_if_exists(filename)

# Function to create directories if they don't exist
def create_directories(colors):
    for color in colors:
        os.makedirs(os.path.join(train_path, color), exist_ok=True)
        os.makedirs(os.path.join(val_path, color), exist_ok=True)

def main():
    # Load the Excel file
    file_path = 'linksTestResultsGroundTruth3.xlsx'  # Replace with your Excel file path
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active  # Or wb['SheetName'] if you want to specify the sheet

    # Extract videos and their colors
    videos_by_color = {'white': [], 'yellow': [], 'red': []}

    for row in sheet.iter_rows(min_row=2):  # Assuming headers are in the first row
        video_url = row[0].value  # Assuming video URLs are in the first column
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
            try:
                print("current url: ", video_url)
                save_path = os.path.join(train_path, color, f"video{idx+1}.avi")
                print("save_path: ", save_path)
                # Call the download function
                download_loom_video_to_file(url_str=video_url, trimmed_filename=save_path)
            except: pass
        
        for idx, video_url in enumerate(val_videos):
            try:
                print("current url: ", video_url)
                save_path = os.path.join(val_path, color, f"video{idx+1}.avi")
                # Call the download function
                download_loom_video_to_file(url_str=video_url, trimmed_filename=save_path)
            except: pass

if __name__ == "__main__":
    main()
