import numpy as np
from PIL import Image

def test():
    # Mock text overlay (720x400)
    text_overlay = Image.new("RGBA", (720, 400), (0, 0, 0, 128))
    text_np = np.array(text_overlay)
    text_rgb = text_np[:, :, :3]
    text_alpha = (text_np[:, :, 3:4] / 255.0)

    # Mock cropped frame (1280x720)
    cropped = np.zeros((1280, 720, 3), dtype=np.uint8)

    text_y = 200
    
    roi = cropped[text_y:text_y+text_overlay.height, :]
    cropped[text_y:text_y+text_overlay.height, :] = roi * (1.0 - text_alpha) + text_rgb * text_alpha

    print("Numpy blending test passed!")

if __name__ == "__main__":
    test()
