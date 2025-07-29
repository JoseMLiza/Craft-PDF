# 📄 Craft-PDF — PDF Generator in VB6

**Craft-PDF** is a class written in Visual Basic 6 that allows you to generate PDF files from scratch, without relying on external libraries. It is designed to be lightweight, easy to integrate, and fully functional within the VB6 environment, generating PDF files compatible with the 1.4 specification.


## 🚀 Features

- ✅ Direct PDF file generation in memory or to disk
- ✅ Support for multiple pages
- ✅ Text output using standard PDF fonts (Helvetica, Times, Courier)
- ✅ Drawing lines, rectangles
- ✅ PNG image insertion with transparency support (via GDI+)
- ✅ Optional compression of stream content using `zlib.dll` or `zlibwapi.dll`
- ✅ Support for reusable `XObject Form` objects and watermarks
- ✅ Text justification and line spacing control
- ✅ Compatible with modern PDF readers

## 📦 Requirements

- Visual Basic 6 (VB6)
- GDI+ (for PNG, BMP images with alpha transparency)
- **Optional:** `zlib.dll` or `zlibwapi.dll` for stream compression (FlateDecode)

## 🧩 Files

- `cCraft.cls` — Main class containing all PDF generation logic

## 📐 Basic Usage

```vb
Dim pdf As New cCraft

Call pdf.StartDoc
Call pdf.SetFont(HELVETICA, 12, vbBlue)
Call pdf.DrawText("Hello world from Craft-PDF", 50, 750)

'-- Save to disk
Call pdf.Save("C:\MyFile.pdf")

'-- Save to memory
Dim Buff As String
Buff = pdf.Contents
```
## 🖼️ Images
The class allows you to embed PNG images and reuse them as XObjects, ideal for logos and watermarks that appear on multiple pages.

```vb
Call pdf.AddImage("logo.png", "Img1") 
Call pdf.DrawImage("Img1", 100, 100, 200, 100)
```

## 🗜️ Compression (Optional)
Craft-PDF can optionally compress PDF stream contents (such as text and image objects) using the FlateDecode filter. This requires one of the following dynamic libraries to be available:

- zlib.dll (standard)
- zlibwapi.dll (alternative build)

Compression reduces file size and improves loading performance in PDF readers.

## Technical Notes
- The generated format is PDF 1.4
- The class manually constructs the PDF structure: header, object bodies, cross-references, and trailer
- No external dependencies are required except GDI+ for PNG transparency

## 📘 License
This project is free to use. You may modify, use, and adapt it to suit your needs.

 ![ITypeComp::Bind](/res/sc_00.png)
