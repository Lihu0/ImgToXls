"use strict";

const imageInput = document.getElementById("imageInput");
const imagePreview = document.getElementById("imagePreview");
const canvas = document.getElementById("canvas");
const context = canvas.getContext("2d");
const img = new Image();
const maxSize = 200;

function numberToExcelColumn(num) {
  let column = "";
  while (num > 0) {
    num--;
    column = String.fromCharCode((num % 26) + 65) + column;
    num = Math.floor(num / 26);
  }
  return column;
}

function rgbToHex(r, g, b) {
  r = Math.min(255, Math.max(0, r));
  g = Math.min(255, Math.max(0, g));
  b = Math.min(255, Math.max(0, b));
  return ((1 << 24) | (r << 16) | (g << 8) | b)
    .toString(16)
    .slice(1)
    .toUpperCase();
}

function resizeImage() {
  let { width, height } = img;
  const scalingFactor = Math.min(maxSize / width, maxSize / height, 1);
  width *= scalingFactor;
  height *= scalingFactor;
  canvas.width = width;
  canvas.height = height;
  context.drawImage(img, 0, 0, width, height);
}

imageInput.addEventListener("change", function (event) {
  const file = event.target.files[0];
  if (file) {
    const fileURL = URL.createObjectURL(file);
    imagePreview.src = fileURL;
    img.src = fileURL;
  }
});

img.onload = function () {
  if (img.width === 0 || img.height === 0) {
    alert("Failed to load the image.");
    return;
  }
  resizeImage();
};

document.getElementById("download").addEventListener("click", function () {
  if (
    canvas.width === 0 ||
    canvas.height === 0 ||
    img.width === 0 ||
    img.height === 0
  ) {
    alert("Please upload an image first.");
    return;
  }
  const ws = {};
  const imageData = context.getImageData(0, 0, canvas.width, canvas.height);
  const data = imageData.data;
  let refStartColumn = "A";
  let refEndColumn = "A";
  let refStartRow = 1;
  let refEndRow = 1;
  for (let i = 0; i < canvas.width; i++) {
    for (let j = 0; j < canvas.height; j++) {
      const index = (j * canvas.width + i) * 4;
      const red = data[index];
      const green = data[index + 1];
      const blue = data[index + 2];
      const cellRed = numberToExcelColumn(i * 3 + 1) + (j + 1);
      ws[cellRed] = {
        t: "n",
        s: { fill: { fgColor: { rgb: rgbToHex(red, 0, 0) } } },
        v: red,
      };
      const cellGreen = numberToExcelColumn(i * 3 + 2) + (j + 1);
      ws[cellGreen] = {
        t: "n",
        s: { fill: { fgColor: { rgb: rgbToHex(0, green, 0) } } },
        v: green,
      };
      const cellBlue = numberToExcelColumn(i * 3 + 3) + (j + 1);
      ws[cellBlue] = {
        t: "n",
        s: { fill: { fgColor: { rgb: rgbToHex(0, 0, blue) } } },
        v: blue,
      };
      const lastColumn = i * 3 + 3;
      if (lastColumn > refEndColumn.charCodeAt(0) - 64) {
        refEndColumn = numberToExcelColumn(lastColumn);
      }
      refEndRow = j + 1;
    }
  }
  ws["!ref"] = `${refStartColumn}${refStartRow}:${refEndColumn}${refEndRow}`;
  const numColumns = canvas.width * 3;
  ws["!cols"] = Array(numColumns).fill({ wch: 2.75 });
  const numRows = canvas.height;
  ws["!rows"] = Array(numRows).fill({ hpt: 50 });
  const wb = { SheetNames: ["Sheet1"], Sheets: { Sheet1: ws } };
  XLSX.writeFile(wb, "image.xlsx");
});
