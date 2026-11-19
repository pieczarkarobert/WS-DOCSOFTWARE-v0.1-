# DocSoftware

A lightweight desktop text and code editor written in Python with `tkinter`. I started building it because heavy word processors feel slow when all you want to do is mix technical notes with a snippet of code, while code editors can be clunky for plain prose and quick tables.

This is an early prototype (**v0.1**) that serves as the foundation for **WS DocStudio**.

## Features

- **Text formatting**: bold, italic, underline, font family, font size, and text color.
- **Tables**: insert a configurable row-and-column grid directly into the document so you can edit cells in place and export cleanly.
- **Code Mode**: switch to a monospace font, insert 4-space indents with `Tab`, drop in code blocks, and get basic syntax highlighting for Python, HTML, and C++.
- **Saving and export**:
  - Documents: `.txt`, `.md`, `.html`, `.rtf`, plus `.docx` and `.pdf` when optional packages are installed.
  - Source code and scripts (UTF-8): `.py`, `.cpp`, `.js`, `.json`, `.bat`, `.sh`.

## Known Limitations (Early v0.1)

No pretending this is a finished suite yet. Right now it still does not support:
- drag-and-drop image placement and rich media embeds,
- live spell checking,
- cloud sync or real-time collaboration,
- macros and scripting automation.

## Getting Started

DocSoftware runs out of the box using Python's built-in standard library (3.9+):

```bash
python main.py
