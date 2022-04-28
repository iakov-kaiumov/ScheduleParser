import os
from PIL import Image
import PIL


FILE_DIR = '/Users/iakov/Documents/Phystech.Center/Графика/photo_compress/'
fixed_height = 1600


def main():
    folders = [FILE_DIR + dr + '/' for dr in os.listdir(FILE_DIR)]
    for folder in folders:

        if not os.path.isdir(folder):
            continue

        files = [folder + dr for dr in os.listdir(folder)]
        for path in files:
            print(path)
            if path.split('.')[-1] != 'jpg':
                continue

            image = Image.open(path)
            height_percent = (fixed_height / float(image.size[1]))
            width_size = int((float(image.size[0]) * float(height_percent)))
            image = image.resize((width_size, fixed_height), PIL.Image.NEAREST)
            image.save(path)


if __name__ == '__main__':
    main()
