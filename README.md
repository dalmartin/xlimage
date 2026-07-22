# xlimage

A **Python** library for extracting **images** from **Excel** files.

## Overview

Microsoft 365 introduced the **Place in Cell** feature, allowing images to be stored directly within worksheet cells rather than as floating drawing objects. While several Python libraries support traditional Excel images, there is currently no library that parses the xml to get these images stored within cells. This is a lightweight library that allows users to extract such images by providing only a sheet and cell. It's possible that more features will be added to this as they come up, feel free to submit an issue if you find a bug or have a suggestion going forward.

With no dependencies (pure stdlib), xlimage is extremely lightweight. This project started due to a lack of maintained python-excel libraries, so it doesn't rely on them to work. `xlimage` aims to provide a simple API for locating, extracting, and working with images stored in excel files through Python.

## Install
```bash
pip install xlimage

```
## Use
To use xlimage, install it to your environment (see above) and import the ImageLoader class. Below is a quick guide on initializing an ImageLoader object and using its methods.

- ImageLoader(workbook_path: str)
    - Constructor for an ImageLoader that takes the workbook at the provided path and runs it through a parser to map cells with images to image filepaths.
    - `workbook_path`: a string that represents the file path for the `workbook` to be parsed. All subsequent functino calls work off of this `workbook`.

- hasImage(sheet: str, cell: str) -> bool
    - has_image will return True if the given cell contains an image, and False otherwise.
    - `sheet`: string that represents the target worksheet's name (should be in the `workbook` provided on initialization) e.g. "sheet1"
    - `cell`: Should be a string that represents the target `cell` in `sheet`

- getImage(sheet: str, cell: str) -> bytes
    - get_image will raise an exception if there is no image in the given cell. If there is an image, it will return that image as a bytes object.
    - The image bytes can be easily converted into a Pillow image, for example, written to filesystem, etc.
    - `sheet`: string that represents the target worksheet's name (should be in the `workbook` provided on initialization) e.g. "sheet1"
    - `cell`: Should be a string that represents the target `cell` in `sheet`

## Example
```python
from xlimage import ImageLoader

# Initialize an ImageLoader. This parses the given workbook, and maps cells to image paths.
xli = ImageLoader("workbook.xlsx")

# Check if an image exists at A1 in sheet1
check = xli.hasImage("sheet1", "A1")

if check:

    # Get an image represented as bytes with .get_image()
    image = xli.getImage("sheet1", "A1")

# Save the image, use Pillow to work with it, or anything else.

```

## Current Status
**Working**
This is still in development, and will likely be an ongoing project if there is significant interest. As of now, the base functionality is working.

## Contributing
Contributions, sample workbooks, bug reports, and format analysis are welcome.
Please submit an issue if you come across a bug or new feature that you'd like to see implemented

## License
GNU GENERAL PUBLIC LICENSE v3.0
