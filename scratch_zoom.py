import os
from PIL import Image
import numpy as np
from moviepy.editor import ImageClip

def test_zoom():
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)
    
    # Create a dummy image
    img = Image.new('RGB', (720, 1280), color = (73, 109, 137))
    img_path = os.path.join(output_dir, 'test_zoom.jpg')
    img.save(img_path)
    
    np_img = np.array(img)
    duration = 3.0
    clip = ImageClip(np_img).set_duration(duration)
    
    def zoom(t):
        return 1.0 + 0.1 * (t / duration)
        
    clip = clip.resize(zoom)
    clip = clip.crop(x_center=360, y_center=640, width=720, height=1280)
    
    clip.write_videofile(os.path.join(output_dir, "zoom_test.mp4"), fps=24)
    print("Zoom test passed!")

if __name__ == "__main__":
    test_zoom()
