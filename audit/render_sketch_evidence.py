"""Lossless encoding of pictures saved by the actual VB6 PictureBox, not a redraw."""
from pathlib import Path
from PIL import Image
import sys
p=Path(sys.argv[1])
for bmp in p.glob('*.bmp'):
    with Image.open(bmp) as im: im.save(bmp.with_suffix('.png'))
print('Converted native VB6 pictures to PNG; inspect PNGs separately for visual QA.')
