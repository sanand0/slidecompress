# slidecompress

A command-line tool that reduces PowerPoint file sizes by intelligently compressing embedded images while maintaining visual quality.

## Features

- Automatically resizes large images to reasonable dimensions (1000x1000 max)
- Optimizes PNG files by reducing colors to 255 while maintaining quality
- Compresses JPEG files with balanced quality settings (75% quality)
- Converts TIF/TIFF/EMF files to optimized PNGs
- Preserves original files by creating a new `.compressed.pptx` file
- Shows detailed compression statistics for each image
- Reports which slides contain the largest remaining images

## Requirements

- Python 3.6+
- ImageMagick (`magick` command must be available in PATH)

## Usage

First, [compress all images in PowerPoint](https://support.microsoft.com/en-us/office/reduce-the-file-size-of-your-powerpoint-presentations-9548ffd4-d853-41e7-8e40-b606bca036b4).

Then, run the tool:

```bash
uvx slidecompress presentation.pptx
```

Or with custom maximum image width:

```bash
uvx slidecompress --width 480 presentation.pptx
```

The compressed file will be saved as `presentation.compressed.pptx` in the same directory.

To overwrite the original file, use the `-f` flag:

```bash
uvx slidecompress -f presentation.pptx
```

## Example Output

```
     image1.png:      9.7KB compressed by      1.4KB
     image2.png:      4.5KB
     image3.jpeg:    19.6KB compressed by      2.3KB
     image4.emf:     98.2KB compressed by     25.3KB and converted to .png
Largest media after compression:
      image4.emf:     72.9KB in slide 4
     image3.jpeg:     17.3KB in slide 3
      image1.png:      8.3KB in slide 1
      image2.png:      4.5KB in slide 2
```

## Build

To build and upload to PyPI:

1. Update the version in `pyproject.toml`
2. Build and upload the package

```
rm -rf dist/
uvx --from build pyproject-build
uvx twine upload -u __token__ dist/*
```

## License

MIT
