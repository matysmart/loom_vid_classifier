### Instructions on how to generate the dataset
### Add dataset to links.xlsx (0-bad, 1-good)
### Edit the "offset" variable so as not to replace existing training dataset
### Output will be in "bad" and "good" folders
### Transfer these to train/val/test
### Train colab notebook is in: loom_video_error_classifier_train.ipynb
### Inference colab notebook is in: loom_video_error_classifier.ipynb (upload linksTestOnly.xlsx)
import argparse
import json
import urllib.request
from moviepy.video.io.VideoFileClip import VideoFileClip
import os
import pandas as pd

offset = 118

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
    # trimmed_filename = f"{id}.avi"  # Final output filename
    print(f"Trimming video {filename} to {trimmed_filename}")
    trim_video(filename, trimmed_filename)

    # Example usage
    delete_file_if_exists(filename)

# Function to remove the query part from the URL
def remove_query_from_url(url):
    return url.split('?')[0]

def main():
    # Read the Excel file
    df = pd.read_excel('links.xlsx')
    df['video'] = df['video'].apply(remove_query_from_url)

    # Create the directory structure if it doesn't exist
    os.makedirs('loom_dataset/bad', exist_ok=True)
    os.makedirs('loom_dataset/good', exist_ok=True)
    
    # Iterate over each row in the dataframe
    for index, row in df.iterrows():
        url = row['video']
        label = row['label']
        
        # Determine save directory based on the label
        if label == 0:
            save_dir = 'loom_dataset\\bad'
            filename = f'video{index+1+offset}.avi'
        elif label == 1:
            save_dir = 'loom_dataset\\good'
            filename = f'video{index+1+offset}.avi'
        
        # Full path to save the video
        save_path = os.path.join(save_dir, filename)
        print(f"save_path: {save_path}")
        # Call the download function
        download_loom_video_to_file(url_str=url, trimmed_filename=save_path)

if __name__ == "__main__":
    main()
