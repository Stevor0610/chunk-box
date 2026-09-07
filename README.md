# Chunk Box

A simple browser-based file downloader that downloads files in chunks and reconstructs them locally.

## Features

* 📦 Download large files in multiple chunks
* 📊 Real-time download progress based on actual bytes downloaded
* 📁 Displays file size and number of chunks
* ⚡ Downloads chunks sequentially
* 🧩 Reconstructs the original file directly in the browser
* 💾 Automatically saves the reconstructed file
* 🌙 Dark-themed interface
* 🚫 No backend server required

## How It Works

Chunk Box stores each file as a collection of smaller chunk files.

The repository contains a `files.json` file that describes the available files:

```json
[
    {
        "name": "LINEPortable.zip",
        "size": 79426777,
        "chunks": 2
    },
    {
        "name": "python-3.14.7-embed-amd64.zip",
        "size": 12673227,
        "chunks": 1
    }
]
```

When a user selects a file, Chunk Box:

1. Reads the file information from `files.json`
2. Downloads each chunk sequentially
3. Tracks the actual number of bytes downloaded
4. Updates the progress bar and percentage
5. Combines all chunks into a `Blob`
6. Creates a local download and saves the original filename

The files themselves are served directly from the repository through GitHub's raw content service.

## Repository Structure

```text
chunk-box/
├── index.html
├── files.json
├── makechunks.yml
└── <file>/
    ├── chunk_part_1
    ├── chunk_part_2
    └── ...
```

For example:

```text
LINEPortable.zip/
├── chunk_part_1
└── chunk_part_2
```

## Using Chunk Box

Open the GitHub Pages site:

**https://stevor0610.github.io/chunk-box/**

Select a file from the list and click **Download**.

The browser will download the required chunks and reconstruct the original file automatically.

## Generating Chunks

The repository uses the `makechunks.yml` GitHub Actions workflow to generate and update chunk files.

The workflow also updates `files.json` with metadata such as:

* Original filename
* Original file size
* Number of chunks

This allows the web interface to discover available files without requiring a separate database or backend service.

## Limitations

Because the file is reconstructed in the browser:

* The browser must have enough available memory to hold the downloaded data.
* The complete file is assembled into a `Blob` before the final download.
* Chunks are currently downloaded sequentially rather than in parallel.

For very large files, browser memory limits may become the main constraint.

## Technology

* HTML
* CSS
* JavaScript
* GitHub Actions
* GitHub Pages
* GitHub raw content

No external JavaScript libraries are required.

## License

See the repository for license information.
