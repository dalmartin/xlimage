# xlimage

A Python library for extracting **Excel "Place in Cell" images** from `.xlsx` files.

## Overview

Microsoft 365 introduced the **Place in Cell** feature, allowing images to be stored directly within worksheet cells rather than as floating drawing objects. While several Python libraries support traditional Excel images, there is currently no library that parses the xml to get these images. This is a lightweight library that allows users to extract such images by providing only a sheet and cell. It's possible that more features will be added to this as they come up, feel free to submit an issue if you have a suggestion or have a suggestion going forward.

`xlimage` aims to provide a simple API for locating, extracting, and working with images stored in excel files through Python.

## API

```python
from xlimage import ImageLoader

xli = ImageLoader("workbook.xlsx")

# Get an image represented as bytes with .get_image()
image = xli.get_image("Sheet1", "A1")

# Save the image, use Pillow to work with it, or anything else.

```

## Current Status
**Working**
This is still in development, and will likely be an ongoing project if there is significant interest. As of now, the base functionality is working.

#### Roadmap for now 
- Unit tests
- PyPI package release
- Documentation and examples

## Contributing
Contributions, sample workbooks, bug reports, and format analysis are welcome.

## License
GNU GENERAL PUBLIC LICENSE v3.0
