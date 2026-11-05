"""Optional visual check after rendering ui-conditional-print's PDF at 96 DPI.

Requires Pillow. Pass PNG and metrics paths or use the repository defaults.
This checks rendered PDF pixels; it does not claim Microsoft Excel parity.
"""
import hashlib
import json
import sys
from pathlib import Path
from PIL import Image

png = Path(sys.argv[1] if len(sys.argv) > 1 else 'tmp/pdfs/conditional-visual-print.png')
metrics = Path(sys.argv[2] if len(sys.argv) > 2 else 'test-results/conditional-print-metrics.json')
image = Image.open(png).convert('RGB')
cells = json.loads(metrics.read_text())['print']
samples = []


def pixel(ref, fraction):
    rect = cells[ref]
    point = (round(rect['x'] + rect['width'] * fraction),
             round(rect['y'] + rect['height'] * .5))
    rgb = image.getpixel(point)
    samples.append({'ref': ref, 'fraction': fraction, 'point': point, 'rgb': rgb})
    return rgb


for ref, expected in {'B2': (189, 215, 238), 'B3': (222, 235, 247),
                      'B5': (255, 255, 255), 'C2': (255, 0, 0),
                      'C3': (255, 255, 255), 'C5': (0, 255, 0),
                      'C9': (128, 255, 128)}.items():
    actual = pixel(ref, .25)
    assert all(abs(a - b) <= 3 for a, b in zip(actual, expected)), (ref, actual, expected)

for ref, fraction in [('C7', .75), ('C8', .75), ('D2', .25), ('D3', .75),
                      ('D6', .75), ('D7', .75), ('D8', .75), ('D9', .75)]:
    actual = pixel(ref, fraction)
    assert min(actual) >= 250, (ref, fraction, actual, 'expected an unpainted white area')

for ref, fraction in [('D2', .05), ('D3', .25), ('D5', .75), ('D9', .25)]:
    actual = pixel(ref, fraction)
    assert actual[2] > actual[0] + 20, (ref, fraction, actual, 'expected the blue gradient')

result = {'status': 'passed', 'imageSize': image.size, 'dpi': 96,
          'pngSHA256': hashlib.sha256(png.read_bytes()).hexdigest(), 'samples': samples,
          'scope': 'Rendered Chromium PDF: scale endpoints/midpoint, bar inside/outside, ignored text/Boolean/blank and numeric formulas. Not an Excel pixel comparison.'}
Path('test-results/conditional-print-pixels.json').write_text(json.dumps(result, indent=2))
print(f'PASS: {len(samples)} rendered-PDF pixel checks and one-page visual inspection fixture')
