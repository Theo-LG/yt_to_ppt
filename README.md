# yt_to_ppt
Extract slides from a Youtube Presentation.

### How to Use
Just change the URL in the main.py with the URL from the video you want to extract. 

### Threshold (1 - epsilon)
This program uses skimage.measure to compare successive frames. But because of YouTube, the frames are slightly blurred so you need to use an epsilon to compare frames. This can lead to two issues (not all slides are taken into account or in the other case some frames are taken when they should not)
