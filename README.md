# xlimage

A Python library for extracting **Excel "Place in Cell" images** from `.xlsx` files.

## Overview

Microsoft 365 introduced the **Place in Cell** feature, allowing images to be stored directly within worksheet cells rather than as floating drawing objects. While several Python libraries support traditional Excel images, support for Place in Cell images is currently limited.

`xlimage` aims to provide a simple API for locating, extracting, and working with these images in Python.

## Goals

- Read Place in Cell images from `.xlsx` files
- Map worksheet cells to their corresponding images
- Return images as Pillow (`PIL.Image`) objects
- Support multiple worksheets
- Expose image metadata and relationships
- Be lightweight and dependency-friendly

## Planned API

```python
from xlimage import CellImageReader

reader = CellImageReader("workbook.xlsx")

image = reader.get_image("Sheet1", "A1")

image.show()
```

## Current Status
#### Work in Progress
This project is actively being reverse-engineered and developed. The XML structures used by Excel's Place in Cell feature are still being documented and tested across workbook versions.
Development Roadmap

#### To-Do for v1.0:
- Extract images from /xl/media
- Pillow integration
- Unit tests
- PyPI package release
- Documentation and examples

## Contributing
Contributions, sample workbooks, bug reports, and format analysis are welcome.

## License
GNU GENERAL PUBLIC LICENSE v3.0
