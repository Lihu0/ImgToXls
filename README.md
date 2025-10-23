# ImgToXls

Converts an image into an Excel sheet, where each pixel's **Red**, **Green**, and **Blue** values are displayed vertically in the same column, like this:

| Pixel 1 | Pixel 2 | Pixel 3 | Pixel 4 |
|---------|---------|---------|---------|
| Red     | Red     | Red     | Red     |
| Green   | Green   | Green   | Green   |
| Blue    | Blue    | Blue    | Blue    |

## Features
- Reads pixel data from any image.
- Generates a fully formatted Excel file with RGB values.
- Works in the browser with **VanillaJS**.

## Frameworks Used
- [SheetJS](https://www.npmjs.com/package/xlsx) – for creating and manipulating Excel files  
- [xlsx-js-style](https://www.npmjs.com/package/xlsx-js-style) – for styling Excel cells  

## Running locally

1. **Clone the repository**:

```bash
git clone https://github.com/Lihu0/ImgToXls.git
cd ImgToXls
```

2. **Open the project**:
Simply open [`index.html`](index.html) in your preferred web browser.

3. **Use the application**:
   - Select an image using the file input.
   - Click **Export to Excel** to generate an Excel file with each pixel's RGB values.
