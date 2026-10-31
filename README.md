![Banner](banner.png)

# dsan-metadata-stripper
A command-line utility that removes potentially revealing metadata from common file formats (PDF, DOCX, JPG, JPEG, and PNG).

For DOCX files, all core properties are cleared (author, title, subject, keywords, comments, last modified by, category, created, modified, revision, identifier, and version), the custom properties part is removed, and the comments part and its references are removed. For images, the picture is rebuilt from its pixels, which drops EXIF data and text chunks.

## Usage
`python main.py <input_path> [--output <output_dir>] [--recursive] [--verbose]`

- `input_path`: Path to a single file or a directory of files.
- `--output`: Output directory. Defaults to `sanitized_output`.
- `--recursive`: Process subdirectories when the input is a directory.

## Behaviour notes
- A run is refused at exit 1 when the output path resolves to the input file, or when the output directory resolves to the input directory.
- In directory mode the output mirrors the input subdirectory structure, so files with the same name in different subdirectories do not overwrite each other.
- The exit code is 1 when the input path does not exist, when a file has an unsupported type, when any file fails to process, or when a directory run processes zero files. Exit 0 means every file was stripped and written.

## License
Copyright (c) ShadowStrikeHQ
