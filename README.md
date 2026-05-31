# Zeta: High-Performance Local File & Text Search Engine

Zeta is a fast, multi-threaded local search utility designed to bypass the limitations of native OS file searching. Built with Python and Tkinter, it provides rapid indexing and deep text searching within files using advanced parallel processing.

## Features
* **Ultra-Fast Search:** Utilizes multi-core processing and memory mapping (`mmap`) to scan files in parallel.
* **Smart Text Matching:** Character normalization for accent-insensitive and case-insensitive matching.
* **Live Preview:** View file contents and highlight matching keywords instantly without freezing the UI.
* **Clean GUI:** Simple and responsive desktop interface built with Tkinter.

## Architecture
Zeta uses a decoupled architecture where the processing core handles heavy disk I/O in worker processes, communicating with the main GUI thread via thread-safe queues. This ensures zero interface lag during intense search operations.

## License
This project is licensed under the MIT License.
