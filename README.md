## PowerPoint size optimizer
This tool reduces the size of a PowerPoint presentation in 2 ways:

* Compress embedded PNGs by using [OptiPNG](http://optipng.sourceforge.net/) (loss-less transformation).
* Convert PNGs to JPEGs by using [ImageMagick](https://imagemagick.org/) (lossy transformation).

### Usage
If you want to use OptiPNG, then you can do:
```
./powerpoint_optimize.py --optipng input.pptx output.pptx
```

If you want to convert PNGs to JPEGs, then you can do:
```
./powerpoint_optimize.py --png-to-jpg input.pptx output.pptx
```
